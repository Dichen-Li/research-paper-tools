"""真实 PptxGenJS 导出测试；依赖可用时执行，不伪造 Office 文件。

可设置 PAPER_REDRAW_NODE 和 PAPER_REDRAW_NODE_MODULES 指向当前运行时。
"""
import json
import base64
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def junction_command(link, target, powershell):
    # 单引号字面量转义，再按 PowerShell 的 UTF-16LE 协议传递；避免 cmd.exe
    # 对含空格、中文、&、%、引号的路径作二次展开，不修改执行策略。
    literal = lambda value: "'" + str(value).replace("'", "''") + "'"
    script = ("$ErrorActionPreference='Stop'; New-Item -ItemType Junction -Path "
              + literal(link) + " -Value " + literal(target) + " | Out-Null")
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    return [powershell, "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded]


class TemplateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.node = os.environ.get("PAPER_REDRAW_NODE") or shutil.which("node")
        cls.packages = os.environ.get("PAPER_REDRAW_NODE_MODULES")
        if not cls.node or not cls.packages or not (Path(cls.packages) / "pptxgenjs").exists():
            raise unittest.SkipTest("Set the actual Node and package directory for integration tests")

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name) / "build 空格"
        self.folder.mkdir()
        shutil.copy2(ROOT / "assets/pptxgenjs-native-template.mjs", self.folder / "build.mjs")
        link = self.folder / "node_modules"
        if sys.platform == "win32":
            powershell = shutil.which("powershell.exe") or shutil.which("pwsh.exe")
            if not powershell:
                self.skipTest("PowerShell unavailable for Windows test-directory junction")
            subprocess.run(junction_command(link, self.packages, powershell), check=True,
                           capture_output=True, timeout=20)
            self.addCleanup(os.rmdir, link)  # 只移除 junction，再清理临时目录。
        else:
            link.symlink_to(self.packages, target_is_directory=True)

    def build(self, data=None, extra=None):
        command = [self.node, str(self.folder / "build.mjs")]
        if data is not None:
            (self.folder / "map.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            command += ["--map", "map.json"]
        result = subprocess.run(command + (extra or []), cwd=ROOT, capture_output=True,
                                text=True, encoding="utf-8", errors="replace", timeout=30)
        return result

    def slide(self):
        with zipfile.ZipFile(self.folder / "editable.pptx") as archive:
            return ET.fromstring(archive.read("ppt/slides/slide1.xml"))

    def test_default_nodes_are_single_objects_and_bend_is_one_path(self):
        result = self.build()
        self.assertEqual(result.returncode, 0, result.stderr)
        root = self.slide()
        named = {sp.find("p:nvSpPr/p:cNvPr", NS).get("name"): sp
                 for sp in root.findall(".//p:sp", NS)}
        self.assertEqual(len([name for name in named if name.startswith("Node:")]), 3)
        self.assertFalse(any(name.startswith("Node-label:") for name in named))
        self.assertTrue(all(named[name].find("p:txBody", NS) is not None
                            for name in named if name.startswith("Node:")))
        edge = named["Edge:question-to-mechanism"]
        self.assertEqual(len(edge.findall("p:spPr/a:custGeom/a:pathLst/a:path", NS)), 1)
        self.assertEqual(edge.find("p:spPr/a:ln/a:tailEnd", NS).get("type"), "triangle")
        other = named["Edge:feedback"]
        for end in ("headEnd", "tailEnd"):
            self.assertEqual(other.find("p:spPr/a:ln/a:" + end, NS).get("type"), "triangle")

    def test_non169_circle_and_fixed_canvas_keep_equal_scales(self):
        for slide in (None, {"width": 12, "height": 7.5}):
            data = {"source": {"width": 1000, "height": 400},
                    "nodes": [{"id": "circle", "type": "ellipse", "text": "X",
                               "box": {"x": 100, "y": 100, "w": 100, "h": 100}}], "edges": []}
            if slide:
                data["slide"] = slide
            result = self.build(data, ["--overwrite"])
            self.assertEqual(result.returncode, 0, result.stderr)
            extent = self.slide().find(".//p:sp/p:spPr/a:xfrm/a:ext", NS)
            self.assertLessEqual(abs(int(extent.get("cx")) - int(extent.get("cy"))), 1)

    def test_reversed_endpoint_map_rejected_before_export(self):
        data = {"source": {"width": 1000, "height": 400}, "nodes": [
            {"id": "a", "type": "rect", "text": "A", "box": {"x": 100, "y": 100, "w": 100, "h": 80}},
            {"id": "b", "type": "rect", "text": "B", "box": {"x": 700, "y": 100, "w": 100, "h": 80}}],
            "edges": [{"id": "wrong", "from": "a", "to": "b", "direction": "forward",
                       "points": [{"x": 700, "y": 140}, {"x": 200, "y": 140}]}]}
        result = self.build(data)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("endpoints", result.stderr)
        self.assertFalse((self.folder / "editable.pptx").exists())

    def test_existing_file_not_overwritten(self):
        self.assertEqual(self.build().returncode, 0)
        before = (self.folder / "editable.pptx").read_bytes()
        self.assertNotEqual(self.build().returncode, 0)
        self.assertEqual((self.folder / "editable.pptx").read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
