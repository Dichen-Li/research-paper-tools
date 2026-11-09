#!/usr/bin/env python3
"""检查原生 PPTX/XLSX 的结构与可编辑性线索，不验证转录或视觉相似。

python check_native_office.py result.pptx [--output private-check.json]
PASS/WARN/FAIL；WARN 返回0，需结合源图复核，不自动改写文件。
仅用于默认原生交付；缩略图与可识别的同框公式兼容图单独处理。
"""

import argparse
import json
import posixpath
import re
import sys
import unicodedata
from pathlib import Path
from urllib.parse import unquote
import xml.etree.ElementTree as ET
import zipfile


P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
NUMERIC_TEXT = re.compile(r"^[+−-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$")
SCRIPT_TEXT = re.compile(r"[_^](?:\{[^{}]+\}|[A-Za-z0-9]+)")


def warn(report, code, **details):
    if len(report["warnings"]) < 100:
        report["warnings"].append({"code": code, **details})
    else:
        report["warnings_suppressed"] = report.get("warnings_suppressed", 0) + 1


def resolve_target(owner, target):
    target = unquote(target.split("#", 1)[0])
    return posixpath.normpath(target.lstrip("/") if target.startswith("/")
                             else posixpath.join(posixpath.dirname(owner), target))


def formula_fallbacks(name, root, roots, report):
    """仅识别有真实公式文本、相同几何框、单张图片的兼容分支。"""
    allowed_elements, allowed_media = set(), set()
    folder, filename = posixpath.split(name)
    rels = roots.get(posixpath.join(folder, "_rels", filename + ".rels"))
    targets = {rel.get("Id"): resolve_target(name, rel.get("Target", ""))
               for rel in list(rels) if rel.get("TargetMode") != "External"} if rels is not None else {}

    def frame(element):
        transform = element.find(".//" + P + "spPr/" + A + "xfrm")
        if transform is None:
            return None
        off, extent = transform.find(A + "off"), transform.find(A + "ext")
        return (dict(off.attrib), dict(extent.attrib)) if off is not None and extent is not None else None

    for alt in root.iter(MC + "AlternateContent"):
        fallback = alt.find(MC + "Fallback")
        if fallback is None or len(fallback) != 1 or fallback[0].tag != P + "pic":
            continue
        picture = fallback[0]
        choice = next((c for c in alt.findall(MC + "Choice")
                       if any((t.text or "").strip() for t in c.iter(M + "t"))
                       and any(True for _ in c.iter(M + "oMath"))), None)
        if choice is None or frame(choice) is None or frame(choice) != frame(picture):
            continue
        blips = list(picture.iter(A + "blip"))
        if len(blips) != 1 or not blips[0].get(R + "embed") or blips[0].get(R + "link"):
            continue
        media = targets.get(blips[0].get(R + "embed"), "")
        if not media.startswith("ppt/media/"):
            continue
        allowed_elements.update(id(element) for element in fallback.iter())
        allowed_media.add(media)
        warn(report, "formula_fallback", part=name, media=media,
             message="识别到同框原生公式与兼容图片；需对照确认图片仅为该公式，未验证像素等价。")
    return allowed_elements, allowed_media


def inspect_cells(sheets, roots, report):
    shared = roots.get("xl/sharedStrings.xml")
    strings = ["".join(t.text or "" for t in cell.iter(S + "t"))
               for cell in list(shared)] if shared is not None else []
    for name, root in sheets:
        for cell in root.iter(S + "c"):
            if cell.find(S + "f") is not None:
                continue
            kind = cell.get("t")
            if kind == "s":
                index = int(cell.findtext(S + "v", "-1"))
                if index < 0 or index >= len(strings):
                    raise ValueError(f"Invalid shared-string index in {name}:{cell.get('r')}")
                text = strings[index]
            elif kind == "inlineStr":
                text = "".join(t.text or "" for t in cell.iter(S + "t"))
            elif kind == "str":
                text = cell.findtext(S + "v", "")
            else:
                continue
            normalized = unicodedata.normalize("NFKC", text).strip()
            if NUMERIC_TEXT.fullmatch(normalized):
                warn(report, "numeric_text", part=name, cell=cell.get("r"), value=text[:100],
                     message="外观似数值但存为文本；核对原图及是否为编号，不自动转换。")
            if SCRIPT_TEXT.search(text):
                warn(report, "possible_plain_script", part=name, cell=cell.get("r"), value=text[:100],
                     message="可能用 _/^ 代替上下标，也可能是合法变量名；请按原图复核。")


def inspect(path):
    path = Path(path)
    report = {"file": str(path), "status": "FAIL", "errors": [], "warnings": [], "counts": {},
              "scope": {"checked": ["package", "native-content hints", "picture restrictions"],
                        "not_checked": ["source transcription", "visual fidelity", "scientific validity", "picture provenance"]}}
    errors = report["errors"]
    kind = path.suffix.lower()
    if kind not in (".pptx", ".xlsx"):
        errors.append("Expected .pptx or .xlsx")
        return report
    try:
        with zipfile.ZipFile(path) as archive:
            listing = archive.namelist()
            names = set(listing)
            if len(names) != len(listing):
                errors.append("Duplicate ZIP package parts")
            if archive.testzip():
                raise ValueError("ZIP CRC check failed")
            main = "ppt/presentation.xml" if kind == ".pptx" else "xl/workbook.xml"
            for required in ("[Content_Types].xml", "_rels/.rels", main):
                if required not in names:
                    errors.append(f"Missing package part: {required}")
            roots = {name: ET.fromstring(archive.read(name)) for name in names
                     if name.endswith((".xml", ".rels", ".vml"))}
            # 检查所有内部链接；不提取包，不读取外部链接。
            for name, root in roots.items():
                if not name.endswith(".rels"):
                    continue
                base = "" if name == "_rels/.rels" else name.rsplit("/_rels/", 1)[0]
                for rel in root.findall(REL + "Relationship"):
                    if rel.get("TargetMode") == "External":
                        continue
                    target = unquote(rel.get("Target", "").split("#", 1)[0])
                    resolved = posixpath.normpath(target.lstrip("/") if target.startswith("/")
                                                 else posixpath.join(base, target))
                    if not target or resolved.startswith("../") or resolved not in names:
                        errors.append(f"Broken internal relationship: {name} -> {target}")
            media = [name for name in names if name.startswith(("ppt/media/", "xl/media/"))
                     and not name.endswith("/")]
            image_refs, allowed_media = [], set()
            for name, root in roots.items():
                if not name.startswith(("ppt/", "xl/")):
                    continue
                exempt, parts = formula_fallbacks(name, root, roots, report)
                allowed_media.update(parts)
                if any(id(element) not in exempt and element.tag.rsplit("}", 1)[-1] in
                       {"pic", "blip", "svgBlip", "imagedata", "oleObj", "oleObject"}
                       for element in root.iter()):
                    image_refs.append(name)
            media = [name for name in media if name not in allowed_media]
            if media or image_refs:
                errors.append("Content pictures/embedded objects are not native reconstruction: "
                              + ", ".join(sorted(set(media + image_refs))))

            if kind == ".pptx":
                slides = [(name, root) for name, root in roots.items()
                          if name.startswith("ppt/slides/slide") and name.endswith(".xml")]
                native = sum(len(root.findall(".//" + P + tag)) for _, root in slides
                             for tag in ("sp", "cxnSp", "graphicFrame"))
                report["counts"] = {"slides": len(slides), "native_objects": native}
                if not slides or not native:
                    errors.append("No editable slide objects found")
                for part, root in slides:
                    objects = sum(len(root.findall(".//" + P + tag)) for tag in ("sp", "cxnSp", "graphicFrame"))
                    if objects > 1500:
                        warn(report, "many_native_objects", part=part, count=objects,
                             message="对象过密，检查是否逐像素描图；数量本身不证明违规。")
                    for geometry in root.iter(A + "custGeom"):
                        count = sum(1 for _ in geometry.iter(A + "pt"))
                        if count > 2000:
                            warn(report, "dense_path", part=part, points=count)
                    for text in root.iter(A + "t"):
                        if SCRIPT_TEXT.search(text.text or ""):
                            warn(report, "possible_plain_script", part=part, value=(text.text or "")[:100])
                report["counts"]["script_markers"] = sum(1 for _, root in slides for e in root.iter()
                                                       if e.get("baseline", "0") not in ("0", ""))
            else:
                sheets = [(name, root) for name, root in roots.items()
                          if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")]
                cells = sum(1 for _, root in sheets for cell in root.findall(".//" + S + "c")
                            if any(cell.find(S + tag) is not None for tag in ("v", "is", "f")))
                report["counts"] = {"sheets": len(sheets), "populated_cells": cells}
                if not sheets or not cells:
                    errors.append("No editable cell content found")
                if any(protection.get("sheet", "0") in ("1", "true", "True")
                       for _, root in sheets for protection in root.findall(S + "sheetProtection")):
                    errors.append("Sheet protection prevents default editable delivery")
                inspect_cells(sheets, roots, report)
                report["counts"]["script_markers"] = sum(1 for root in roots.values()
                                                       for e in root.iter(S + "vertAlign")
                                                       if e.get("val") in ("subscript", "superscript"))
    except (OSError, ValueError, ET.ParseError, zipfile.BadZipFile, RuntimeError) as error:
        errors.append(f"Unreadable package: {error}")
    report["status"] = "FAIL" if errors else ("WARN" if report["warnings"] else "PASS")
    return report


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    report = inspect(args.file)
    rendered = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        if args.output.resolve() == args.file.resolve():
            parser.error("Report must not overwrite the Office input")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered, encoding="utf-8")
    print(rendered, end="")
    return 1 if report["status"] == "FAIL" else 0


if __name__ == "__main__":
    raise SystemExit(main())
