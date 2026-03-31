#!/usr/bin/env python3
"""Lightweight VBA module formatter/optimizer for exported .bas/.cls/.frm code.

Usage:
  python tools/optimize_vba_module.py input.bas [output.bas]
"""
from __future__ import annotations
import re
import sys
from pathlib import Path

DECL_RE = re.compile(r'^(Public|Private|Friend)?\s*(Sub|Function|Property)\b', re.I)
COMMENT_BANNER = re.compile(r"^\s*'{3,}.*$")
MULTI_SPACE = re.compile(r"[ \t]{2,}")


def normalize_line(line: str) -> str:
    line = line.rstrip()
    if not line:
        return ""
    # normalize indentation tabs to 4 spaces
    line = line.replace("\t", "    ")
    # trim right and compress repeated spaces outside string literals (simple heuristic)
    if '"' not in line:
        line = MULTI_SPACE.sub(" ", line)
    # standardize control keywords casing (safe/common)
    replacements = {
        "end sub": "End Sub",
        "end function": "End Function",
        "end if": "End If",
        "elseif": "ElseIf",
        "select case": "Select Case",
        "end select": "End Select",
        "on error goto": "On Error GoTo",
        "option explicit": "Option Explicit",
    }
    low = line.lower()
    for k, v in replacements.items():
        if low.strip().startswith(k):
            prefix_len = len(line) - len(line.lstrip())
            line = (" " * prefix_len) + v + line[prefix_len + len(k):]
            break
    return line


def optimize_vba(text: str) -> str:
    lines = text.splitlines()
    out: list[str] = []

    has_option_explicit = any(l.strip().lower() == "option explicit" for l in lines)

    # keep header attributes as-is
    i = 0
    while i < len(lines) and lines[i].lstrip().startswith("Attribute "):
        out.append(lines[i].rstrip())
        i += 1

    if out and out[-1] != "":
        out.append("")

    if not has_option_explicit:
        out.append("Option Explicit")
        out.append("")

    prev_blank = False
    for line in lines[i:]:
        n = normalize_line(line)

        # drop noisy long quote banners like '''''''''''''''''
        if COMMENT_BANNER.match(n):
            continue

        is_blank = (n.strip() == "")
        if is_blank and prev_blank:
            continue

        # ensure one blank line before procedure declarations
        if DECL_RE.match(n.strip()) and out and out[-1].strip() != "":
            out.append("")

        out.append(n)
        prev_blank = is_blank

    return "\n".join(out).rstrip() + "\n"


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python tools/optimize_vba_module.py input.bas [output.bas]")
        return 1

    src = Path(sys.argv[1])
    dst = Path(sys.argv[2]) if len(sys.argv) >= 3 else src

    text = src.read_text(encoding="utf-8", errors="ignore")
    optimized = optimize_vba(text)
    dst.write_text(optimized, encoding="utf-8")
    print(f"Optimized: {src} -> {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
