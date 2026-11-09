#!/usr/bin/env python3
"""Render PPTX/XLSX with isolated LibreOffice and Poppler; QA only.

Adapted from the SciToolsmith renderer; extended for paper-redraw.
Reports the actual renderer and its validation limits. Does not silently fall
back to Quick Look, automate a GUI app, install software, or change the input.
"""

from __future__ import annotations

import argparse
import json
import ntpath
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def office_candidates(platform_name=None, environment=None):
    """单独构造平台候选，便于在非 Windows 主机检查路径分支。"""
    platform_name = sys.platform if platform_name is None else platform_name
    environment = os.environ if environment is None else environment
    if platform_name == "win32":
        candidates = ["soffice.com", "soffice.exe", "libreoffice", "soffice"]
        roots = dict.fromkeys(environment.get(key) for key in
                              ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"))
        for root in roots:
            if root:
                candidates += [ntpath.join(root, "LibreOffice", "program", name)
                               for name in ("soffice.com", "soffice.exe")]
        return candidates
    candidates = ["libreoffice", "soffice"]
    if platform_name == "darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return candidates


def find_command(explicit: str | None, candidates: tuple[str, ...]) -> str:
    if explicit:
        explicit = str(Path(explicit).expanduser())
        resolved = shutil.which(explicit) if not Path(explicit).is_file() else explicit
        if resolved:
            return str(Path(resolved).absolute())
        raise RuntimeError(f"command not found: {explicit}")
    for candidate in candidates:
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
    raise RuntimeError(f"required command not found: {' or '.join(candidates)}")


def run_checked(command: list[str], timeout: int) -> subprocess.CompletedProcess[str]:
    if os.name == "nt" and Path(command[0]).suffix.lower() in (".cmd", ".bat"):
        raise RuntimeError("batch wrapper found; supply the native .exe/.com executable")
    result = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        errors="replace",
        timeout=timeout,
    )
    if result.returncode != 0:
        detail = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part.strip())
        raise RuntimeError(f"command failed ({result.returncode}): {' '.join(command)}\n{detail}")
    return result


def render(
    source: Path,
    output_dir: Path,
    dpi: int,
    keep_pdf: bool,
    timeout: int,
    soffice_arg: str | None,
    pdftoppm_arg: str | None,
    round_number: int | None = None,
) -> dict[str, object]:
    source = source.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() not in (".pptx", ".xlsx"):
        raise RuntimeError(f"input must be an existing .pptx/.xlsx file: {source}")
    if dpi < 72 or dpi > 600:
        raise RuntimeError("dpi must be between 72 and 600")
    if timeout <= 0:
        raise RuntimeError("timeout must be positive")

    output_dir = output_dir.expanduser().resolve()
    if round_number is not None:
        if round_number < 1:
            raise RuntimeError("round must be a positive integer")
        output_dir = output_dir / f"r{round_number}"
    output_dir.mkdir(parents=True, exist_ok=True)
    existing_outputs = sorted(output_dir.glob("page-*.png"))
    if existing_outputs or (output_dir / "rendered.pdf").exists():
        raise RuntimeError(
            "output directory already contains render outputs; use a fresh task directory"
        )
    soffice = find_command(soffice_arg or os.environ.get("SOFFICE"), tuple(office_candidates()))
    pdftoppm = find_command(pdftoppm_arg or os.environ.get("PDFTOPPM"), ("pdftoppm",))
    version_result = run_checked([soffice, "--version"], min(timeout, 15))
    version = (version_result.stdout or version_result.stderr).strip()[:200]
    export_format = "pdf"
    if source.suffix.lower() == ".xlsx":
        # 每个工作表一页，避免宽表分页造成看似丢列；仍需核对实际页数和字号。
        export_format = 'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}'

    with tempfile.TemporaryDirectory(prefix=".sci-lo-profile-", dir=output_dir) as profile_raw:
        with tempfile.TemporaryDirectory(prefix=".sci-render-", dir=output_dir) as work_raw:
            profile = Path(profile_raw).resolve()
            work = Path(work_raw).resolve()
            profile_uri = profile.as_uri()

            run_checked(
                [
                    soffice,
                    f"-env:UserInstallation={profile_uri}",
                    "--headless",
                    "--convert-to",
                    export_format,
                    "--outdir",
                    str(work),
                    str(source),
                ],
                timeout,
            )

            pdf_candidates = sorted(work.glob("*.pdf"))
            if len(pdf_candidates) != 1:
                raise RuntimeError(f"LibreOffice produced {len(pdf_candidates)} PDF files")
            generated_pdf = pdf_candidates[0]

            prefix = output_dir / "page"
            run_checked(
                [pdftoppm, "-png", "-r", str(dpi), str(generated_pdf), str(prefix)],
                timeout,
            )
            pages = sorted(output_dir.glob("page-*.png"),
                           key=lambda p: int(p.stem.rsplit("-", 1)[-1]))
            if not pages:
                raise RuntimeError("pdftoppm produced no slide images")

            kept_pdf: Path | None = None
            if keep_pdf:
                kept_pdf = output_dir / "rendered.pdf"
                shutil.copy2(generated_pdf, kept_pdf)

    return {
        "input": str(source),
        "outputDir": str(output_dir),
        "pages": [str(path) for path in pages],
        "pdf": str(kept_pdf) if kept_pdf else None,
        "dpi": dpi,
        "isolatedLibreOfficeProfile": True,
        "sourceType": source.suffix.lower().lstrip("."),
        "renderer": {"name": "LibreOffice", "version": version, "rasterizer": "pdftoppm"},
        "blind_spots": ["Font substitution and complex formula/shape compatibility may differ from Microsoft Office.",
                        "Rendering does not verify transcription, rich-text semantics, or node binding."],
        "round": round_number,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("office", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--dpi", type=int, default=160)
    parser.add_argument("--keep-pdf", action="store_true")
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--soffice")
    parser.add_argument("--pdftoppm")
    parser.add_argument("--round", type=int, dest="round_number")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        report = render(
            source=args.office,
            output_dir=args.output_dir,
            dpi=args.dpi,
            keep_pdf=args.keep_pdf,
            timeout=args.timeout,
            soffice_arg=args.soffice,
            pdftoppm_arg=args.pdftoppm,
            round_number=args.round_number,
        )
    except (RuntimeError, OSError, subprocess.TimeoutExpired) as error:
        print(json.dumps({"ready": False, "error": str(error), "renderer": "LibreOffice",
                          "verification": "not rendered; do not report visual validation"}, ensure_ascii=True))
        return 1
    print(json.dumps({"ready": True, **report}, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
