#!/usr/bin/env python3
"""Best-effort export + optimization of VBA streams from an .xlsm workbook.

Outputs cleaned module files under ./optimized_vba.
"""
from __future__ import annotations
import re
import sys
import zipfile
from pathlib import Path

THIS_DIR = Path(__file__).resolve().parent
if str(THIS_DIR) not in sys.path:
    sys.path.insert(0, str(THIS_DIR))

from extract_vba import CFB, decompress_vba
from optimize_vba_module import optimize_vba

SKIP_PREFIXES = ("__SRP",)
SKIP_NAMES = {"dir", "_VBA_PROJECT"}
PRINTABLE = set(chr(i) for i in range(32, 127)) | {"\n", "\r", "\t"}


def normalize_raw_text(data: bytes) -> str:
    s = data.decode("latin1", errors="ignore")
    # Replace non-printable characters with line breaks (better token boundary recovery).
    s = "".join(ch if ch in PRINTABLE else "\n" for ch in s)

    # Collapse repeated whitespace/newlines.
    s = re.sub(r"\r\n?", "\n", s)
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)

    # Attempt to join split keyword fragments: "Attribut e" -> "Attribute".
    s = re.sub(r"\b([A-Za-z]{2,}) ([A-Za-z]{1,3})\b", lambda m: m.group(1) + m.group(2), s)

    # Keep only likely code lines to reduce binary noise.
    likely: list[str] = []
    patterns = (
        "Attribute", "Option", "Sub", "Function", "Property", "Dim", "Set ",
        "If ", "Else", "End ", "Select", "Case", "For ", "Next", "Do ",
        "Loop", "With ", "Public", "Private", "Const", "Enum", "Type", "As ",
        "On Error", "MsgBox", "Worksheet", "Workbook", "UserForm", "Call ",
    )
    for raw_line in s.split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        if any(p.lower() in line.lower() for p in patterns):
            likely.append(line)

    return "\n".join(likely) + ("\n" if likely else "")


def main() -> int:
    workbook = Path("SP5_Production_v1.xlsm")
    outdir = Path("optimized_vba")
    outdir.mkdir(exist_ok=True)

    with zipfile.ZipFile(workbook) as zf:
        vba_bin = zf.read("xl/vbaProject.bin")

    cfb = CFB(vba_bin)
    streams = cfb.list_streams_in_storage("VBA")

    written = 0
    for name, idx in streams:
        base = name.split("/")[-1]
        if base in SKIP_NAMES or any(base.startswith(p) for p in SKIP_PREFIXES):
            continue

        raw = cfb.read_stream_by_index(idx)
        dec = decompress_vba(raw)
        text = normalize_raw_text(dec)
        if not text.strip():
            continue

        optimized = optimize_vba(text)
        out_file = outdir / f"{base}.bas"
        out_file.write_text(optimized, encoding="utf-8")
        written += 1

    print(f"Generated {written} optimized module files in {outdir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
