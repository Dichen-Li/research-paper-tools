"""回归结构检查器的真实边界；临时 XML 包仅作单元测试，非交付样例。"""

import importlib.util
from pathlib import Path
import tempfile
import unittest
import zipfile


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("native", ROOT / "scripts/check_native_office.py")
NATIVE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(NATIVE)


class NativeOfficeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def package(self, suffix, extra=None):
        is_ppt = suffix == ".pptx"
        main = "ppt/presentation.xml" if is_ppt else "xl/workbook.xml"
        root_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     f'<Relationship Id="rId1" Target="{main}" Type="officeDocument"/>'
                     '</Relationships>')
        parts = {"[Content_Types].xml": "<Types/>", "_rels/.rels": root_rels, main: "<document/>"}
        if is_ppt:
            parts["ppt/slides/slide1.xml"] = (
                '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                '<p:spTree><p:sp/><p:cxnSp/></p:spTree></p:sld>')
        else:
            parts["xl/worksheets/sheet1.xml"] = (
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row><c r="A1"><v>3.14</v></c></row></sheetData></worksheet>')
        parts.update(extra or {})
        path = self.root / ("result" + suffix)
        with zipfile.ZipFile(path, "w") as archive:
            for name, content in parts.items():
                archive.writestr(name, content)
        return path

    def test_native_slides_and_document_thumbnail(self):
        path = self.package(".pptx", {"docProps/thumbnail.jpeg": b"preview"})
        result = NATIVE.inspect(path)
        self.assertEqual(result["status"], "PASS")
        self.assertEqual(result["counts"]["native_objects"], 2)

    def test_picture_proxy_rejected_even_without_media_bytes(self):
        path = self.package(".pptx", {"ppt/slideMasters/slideMaster1.xml":
            '<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:pic/></p:sldMaster>'})
        self.assertEqual(NATIVE.inspect(path)["status"], "FAIL")

    def test_embedded_svg_not_treated_as_native_geometry(self):
        path = self.package(".pptx", {"ppt/media/image1.svg": "<svg/>"})
        self.assertEqual(NATIVE.inspect(path)["status"], "FAIL")

    def test_editable_cells_pass(self):
        result = NATIVE.inspect(self.package(".xlsx"))
        self.assertEqual(result["status"], "PASS")
        self.assertEqual(result["counts"]["populated_cells"], 1)

    def test_protected_sheet_rejected(self):
        path = self.package(".xlsx", {"xl/worksheets/sheet2.xml":
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetProtection sheet="1"/></worksheet>'})
        self.assertEqual(NATIVE.inspect(path)["status"], "FAIL")

    def test_broken_relationship_rejected(self):
        path = self.package(".pptx", {"ppt/_rels/presentation.xml.rels":
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId2" Target="slides/missing.xml" Type="slide"/>'
            '</Relationships>'})
        self.assertEqual(NATIVE.inspect(path)["status"], "FAIL")

    def test_empty_or_corrupt_file_rejected(self):
        path = self.root / "broken.pptx"
        path.write_bytes(b"not a ZIP")
        self.assertEqual(NATIVE.inspect(path)["status"], "FAIL")

    def test_numeric_text_and_literal_scripts_warn_without_rewriting(self):
        sheet = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                 '<sheetData><row><c r="A1" t="inlineStr"><is><t>0.931</t></is></c>'
                 '<c r="B1" t="inlineStr"><is><t>T_max</t></is></c>'
                 '<c r="C1" t="inlineStr"><is><t>0012</t></is></c></row></sheetData></worksheet>')
        path = self.package(".xlsx", {"xl/worksheets/sheet1.xml": sheet})
        before = path.read_bytes()
        result = NATIVE.inspect(path)
        self.assertEqual(result["status"], "WARN")
        self.assertEqual({w["code"] for w in result["warnings"]},
                         {"numeric_text", "possible_plain_script"})
        self.assertEqual(path.read_bytes(), before)

    def test_dense_native_objects_warn_not_certify_provenance(self):
        xml = ('<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
               '<p:spTree>' + '<p:sp/>' * 1501 + '</p:spTree></p:sld>')
        result = NATIVE.inspect(self.package(".pptx", {"ppt/slides/slide1.xml": xml}))
        self.assertEqual(result["status"], "WARN")
        self.assertIn("many_native_objects", [w["code"] for w in result["warnings"]])
        self.assertEqual(result["warnings"][0]["part"], "ppt/slides/slide1.xml")

    def formula_sample(self, valid=True, outside=False):
        frame = '<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="2000" cy="1000"/></a:xfrm></p:spPr>'
        math = '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>' if valid else '<m:oMath/>'
        picture = '<p:pic>' + frame + '<p:blipFill><a:blip r:embed="rId9"/></p:blipFill></p:pic>'
        slide = ('<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
                 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                 '<p:spTree><mc:AlternateContent><mc:Choice Requires="m"><p:sp>' + frame + math
                 + '</p:sp></mc:Choice><mc:Fallback>' + picture + '</mc:Fallback></mc:AlternateContent>'
                 + (picture if outside else '') + '</p:spTree></p:sld>')
        rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId9" Target="../media/formula.png" Type="image"/></Relationships>')
        return self.package(".pptx", {"ppt/slides/slide1.xml": slide,
                                     "ppt/slides/_rels/slide1.xml.rels": rels,
                                     "ppt/media/formula.png": b"unit-test-only"})

    def test_recognized_formula_fallback_requires_visual_review(self):
        result = NATIVE.inspect(self.formula_sample())
        self.assertEqual(result["status"], "WARN")
        self.assertIn("formula_fallback", [w["code"] for w in result["warnings"]])

    def test_empty_math_does_not_whitelist_picture(self):
        self.assertEqual(NATIVE.inspect(self.formula_sample(valid=False))["status"], "FAIL")

    def test_picture_outside_formula_not_whitelisted(self):
        self.assertEqual(NATIVE.inspect(self.formula_sample(outside=True))["status"], "FAIL")


if __name__ == "__main__":
    unittest.main()
