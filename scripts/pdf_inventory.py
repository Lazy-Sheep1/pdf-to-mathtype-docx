#!/usr/bin/env python3
"""Inventory PDF text/geometry and render local previews; no OCR or upload."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

import fitz


def inventory(source: Path, out: Path, dpi: int = 120) -> dict:
    if not 0 <= dpi <= 600:
        raise ValueError("Preview DPI must be 0 (disabled) or 1..600")
    if out.exists() and any(out.iterdir()):
        raise FileExistsError("Use a new/empty output directory to avoid mixing runs")
    digest = hashlib.sha256()
    with source.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    result = {"schema_version": 1, "source_name": source.name,
              "source_sha256": digest.hexdigest(), "pages": [],
              "notice": "Extraction order is not verified reading order. Text presence does not prove OCR accuracy or complete content."}
    with fitz.open(source) as pdf:
        if pdf.needs_pass:
            raise ValueError("Encrypted PDF requires a separately unlocked authorized copy")
        out.mkdir(parents=True, exist_ok=True)
        for index, page in enumerate(pdf, 1):
            text = page.get_text("text", sort=False)
            # Exclude binary image data; retain font runs, line geometry and text.
            data = page.get_text("dict", flags=fitz.TEXTFLAGS_DICT & ~fitz.TEXT_PRESERVE_IMAGES)
            prefix = f"page-{index:03d}"
            (out / f"{prefix}.txt").write_text(text, encoding="utf-8")
            (out / f"{prefix}.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            info = {"page": index, "width_pt": page.rect.width, "height_pt": page.rect.height,
                    "rotation": page.rotation, "text_characters": len(text.strip()),
                    "needs_ocr_review": not bool(text.strip()),
                    "image_resources": len(page.get_images()),
                    "drawing_paths": len(page.get_drawings()),
                    "fonts": sorted({font[3] for font in page.get_fonts()}),
                    "text_file": f"{prefix}.txt", "geometry_file": f"{prefix}.json"}
            if dpi:
                page.get_pixmap(dpi=dpi, alpha=False, annots=False).save(out / f"{prefix}.png")
                info["preview"] = f"{prefix}.png"
            result["pages"].append(info)
    result["page_count"] = len(result["pages"])
    (out / "inventory.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--dpi", type=int, default=120, help="Preview DPI; 0 disables rendering")
    args = parser.parse_args()
    try:
        result = inventory(args.pdf, args.out, args.dpi)
    except (OSError, ValueError, RuntimeError) as exc:
        parser.exit(2, f"Inventory failed: {exc}\n")
    print(json.dumps({"page_count": result["page_count"], "source_sha256": result["source_sha256"],
                      "pages_without_text": [p["page"] for p in result["pages"] if p["needs_ocr_review"]]}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
