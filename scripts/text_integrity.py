#!/usr/bin/env python3
"""Compare expected prose to explicitly mapped regions in a Word-rendered PDF."""
from __future__ import annotations

import argparse
import difflib
import json
import math
from pathlib import Path
import re
import unicodedata

import fitz

LIGATURES = str.maketrans({"\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl", "\ufb03": "ffi",
                          "\ufb04": "ffl", "\ufb05": "st", "\ufb06": "st"})


def normalize(text: str) -> str:
    # Do not normalize mathematical alphabets, punctuation, digits or real hyphens.
    return re.sub(r"\s+", "", unicodedata.normalize("NFC", text.translate(LIGATURES)).replace("\u00ad", ""))


def remove_approved_breaks(text: str, patterns: list[str]) -> str:
    for pattern in patterns:
        if not isinstance(pattern, str) or not re.fullmatch(r"\w+-\n\w+", pattern):
            raise ValueError("discretionary_breaks must contain exact word-\\ncontinuation patterns")
        if pattern not in text:
            raise ValueError(f"Approved break is not present in rendered region: {pattern!r}")
        text = text.replace(pattern, pattern.replace("-\n", ""))
    return text


def audit(pdf_path: Path, manifest: dict) -> dict:
    entries = manifest.get("entries")
    if not isinstance(entries, list) or not entries:
        raise ValueError("Manifest must contain a nonempty entries list")
    seen = set()
    results = []
    with fitz.open(pdf_path) as pdf:
        if pdf.needs_pass:
            raise ValueError("PDF is encrypted")
        for entry in entries:
            identifier = entry.get("id")
            if not isinstance(identifier, str) or not identifier or identifier in seen:
                raise ValueError("Every entry needs a unique nonempty string id")
            seen.add(identifier)
            if not isinstance(entry.get("text"), str) or not normalize(entry["text"]):
                raise ValueError(f"{identifier}: expected text is empty or invalid")
            regions = entry.get("regions")
            if not isinstance(regions, list) or not regions:
                raise ValueError(f"{identifier}: regions must be nonempty")
            parts = []
            for region in regions:
                page_number, bbox = region.get("page"), region.get("bbox")
                if type(page_number) is not int or not 1 <= page_number <= len(pdf):
                    raise ValueError(f"{identifier}: invalid 1-based page number")
                if (not isinstance(bbox, list) or len(bbox) != 4 or
                        any(type(v) not in (int, float) or not math.isfinite(v) for v in bbox)):
                    raise ValueError(f"{identifier}: bbox must contain four finite numbers")
                page = pdf[page_number - 1]
                if page.rotation:
                    raise ValueError("Normalize rotated pages before mapping text regions")
                rect = fitz.Rect(bbox)
                if rect.is_empty or not page.rect.contains(rect):
                    raise ValueError(f"{identifier}: bbox is empty or outside the page")
                # Region order is supplied by the human/agent, never inferred across columns.
                parts.append(page.get_text("text", clip=rect, sort=True))
            rendered = "\n".join(parts)
            checked = remove_approved_breaks(rendered, entry.get("discretionary_breaks", []))
            expected, actual = normalize(entry["text"]), normalize(checked)
            differences = []
            for tag, i, j, k, l in difflib.SequenceMatcher(None, expected, actual, autojunk=False).get_opcodes():
                if tag != "equal":
                    differences.append({"operation": tag, "expected_offset": i, "actual_offset": k,
                                        "expected": expected[i:j], "actual": actual[k:l]})
            results.append({"id": identifier, "match": expected == actual,
                            "expected_characters": len(expected), "rendered_characters": len(actual),
                            "differences": differences, "rendered_text": rendered})
    return {"schema_version": 1, "checked_entries": len(results),
            "matched_entries": sum(r["match"] for r in results),
            "pass": all(r["match"] for r in results),
            "scope_notice": "Only listed regions are checked. This is not a whole-document coverage, visual overlap, or equation-semantic audit.",
            "entries": results}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()
    if args.out.resolve() in {args.pdf.resolve(), args.manifest.resolve()}:
        parser.error("Output must not overwrite an input")
    try:
        report = audit(args.pdf, json.loads(args.manifest.read_text(encoding="utf-8-sig")))
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    except (OSError, ValueError, RuntimeError, AttributeError, TypeError) as exc:
        parser.exit(2, f"Text audit failed: {exc}\n")
    print(json.dumps({key: report[key] for key in ("checked_entries", "matched_entries", "pass")}))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
