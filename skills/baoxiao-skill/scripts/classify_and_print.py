#!/usr/bin/env python3
"""Classify Chinese invoice files and optionally print them using the user's rules.

Default behavior is dry-run JSON output. Add --print to submit lp jobs.
"""

from __future__ import annotations

import argparse
import json
import subprocess
from pathlib import Path
from typing import Iterable


PDF_EXT = ".pdf"
IMAGE_EXTS = {".jpg", ".jpeg", ".png"}


def pdf_text(path: Path) -> str:
    try:
        from pypdf import PdfReader

        reader = PdfReader(str(path))
        return "\n".join((page.extract_text() or "") for page in reader.pages[:2])
    except Exception as exc:  # pragma: no cover - diagnostic path
        return f"__EXTRACT_ERROR__ {exc}"


def classify_pdf(path: Path) -> str:
    text = pdf_text(path)
    if "增值税专用发票" in text or "专用发票" in text:
        return "vat_special"
    if "普通发票" in text or "增值税电子普通发票" in text:
        return "ordinary"
    return "other_pdf"


def iter_files(paths: Iterable[Path]) -> list[Path]:
    out: list[Path] = []
    for path in paths:
        if path.is_dir():
            out.extend(
                p
                for p in sorted(path.rglob("*"))
                if p.is_file() and p.suffix.lower() in {PDF_EXT, *IMAGE_EXTS}
            )
        elif path.is_file() and path.suffix.lower() in {PDF_EXT, *IMAGE_EXTS}:
            out.append(path)
    return out


def submit_print_job(printer: str, files: list[Path], copies: int) -> dict:
    if not files:
        return {"copies": copies, "file_count": 0, "stdout": "", "stderr": ""}
    cmd = [
        "lp",
        "-d",
        printer,
        "-o",
        "PageSize=A5",
        "-o",
        "media=A5",
        "-o",
        "fit-to-page",
        "-n",
        str(copies),
        *[str(p) for p in files],
    ]
    proc = subprocess.run(cmd, text=True, capture_output=True, check=False)
    return {
        "copies": copies,
        "file_count": len(files),
        "returncode": proc.returncode,
        "stdout": proc.stdout.strip(),
        "stderr": proc.stderr.strip(),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("paths", nargs="+", type=Path, help="Files or directories to classify")
    parser.add_argument("--printer", default=None, help="CUPS printer name. Defaults to system default.")
    parser.add_argument("--print", action="store_true", dest="do_print", help="Submit print jobs")
    args = parser.parse_args()

    groups = {
        "vat_special": [],
        "ordinary": [],
        "other_pdf": [],
        "payment_images": [],
    }

    for path in iter_files(args.paths):
        suffix = path.suffix.lower()
        if suffix == PDF_EXT:
            groups[classify_pdf(path)].append(path)
        elif suffix in IMAGE_EXTS:
            groups["payment_images"].append(path)

    result = {
        "rules": {
            "vat_special": "A5, 2 copies",
            "ordinary": "A5, 1 copy",
            "other_pdf": "A5, 1 copy unless user says otherwise",
            "payment_images": "A5, 1 copy",
        },
        "groups": {k: [str(p) for p in v] for k, v in groups.items()},
        "print_jobs": {},
    }

    if args.do_print:
        printer = args.printer
        if not printer:
            lpstat = subprocess.run(["lpstat", "-d"], text=True, capture_output=True, check=False)
            # macOS Chinese output often ends with the default printer name.
            printer = lpstat.stdout.strip().split()[-1]
        result["printer"] = printer
        result["print_jobs"]["vat_special"] = submit_print_job(printer, groups["vat_special"], 2)
        result["print_jobs"]["ordinary_and_other_pdf"] = submit_print_job(
            printer, groups["ordinary"] + groups["other_pdf"], 1
        )
        result["print_jobs"]["payment_images"] = submit_print_job(printer, groups["payment_images"], 1)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
