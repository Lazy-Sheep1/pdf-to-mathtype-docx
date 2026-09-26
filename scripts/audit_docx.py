#!/usr/bin/env python3
"""Audit MathType OLE packaging and layout risks in a user-supplied DOCX.

Uses the standard library. If olefile is installed, Equation Native streams
are checked as well. This script never starts Word or MathType and never
follows external relationships. Structural success is not an editability or
visual-quality guarantee.
"""
from __future__ import annotations

import argparse
from collections import Counter
from io import BytesIO
import json
from pathlib import Path
import posixpath
import re
import sys
from typing import Any
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O = "urn:schemas-microsoft-com:office:office"
V = "urn:schemas-microsoft-com:vml"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
MATH_TYPE_PROGID = "Equation.DSMT4"
SCOPE_NOTICE = (
    "A passing structural audit does not prove that an equation can be opened "
    "and edited in MathType, that its content is correct, or that Word renders "
    "it without clipping, overlap, or missing glyphs. Verify these in Word and "
    "MathType and inspect a fresh Word-rendered PDF separately."
)


def _load_olefile():
    try:
        import olefile
    except ImportError:
        return None
    return olefile


def _enabled(element: ET.Element | None) -> bool:
    return element is not None and element.get(f"{{{W}}}val", "true").lower() not in {
        "0", "false", "off", "no"
    }


def _rel_part(part: str) -> str:
    parent, name = posixpath.split(part)
    return posixpath.join(parent, "_rels", name + ".rels")


def _target_part(owner: str, target: str) -> str:
    """Resolve a package relationship without accessing a filesystem or URL."""
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise ValueError("Relationship target is not a plain internal package part")
    decoded = unquote(parsed.path)
    if not decoded or "\\" in decoded or "\x00" in decoded:
        raise ValueError("Relationship target is empty or uses invalid path characters")
    if decoded.startswith("/"):
        resolved = posixpath.normpath(decoded.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(owner), decoded))
    if resolved in {"", ".", ".."} or resolved.startswith("../"):
        raise ValueError("Relationship target escapes the package root")
    return resolved


def _xml(data: bytes) -> ET.Element:
    if re.search(br"<!\s*(?:DOCTYPE|ENTITY)\b", data, re.IGNORECASE):
        raise ValueError("DTD and entity declarations are not accepted")
    return ET.fromstring(data)


def _native_stream(payload: bytes, olefile_module: Any) -> dict[str, Any]:
    if olefile_module is None:
        return {"status": "not_checked_dependency_unavailable"}
    try:
        ole = olefile_module.OleFileIO(BytesIO(payload))
        try:
            if not ole.exists("Equation Native"):
                return {"status": "missing"}
            size = ole.get_size("Equation Native")
            stream = ole.openstream("Equation Native")
            try:
                nonempty = bool(stream.read(1))
            finally:
                stream.close()
            return {"status": "nonempty" if size > 0 and nonempty else "empty", "bytes": size}
        finally:
            ole.close()
    except Exception as exc:
        # Do not include exception text: it can contain host paths or content.
        return {"status": "invalid_compound_file", "exception_type": type(exc).__name__}


def audit_docx(path: str | Path, expected_mathtype: int | None = None) -> dict[str, Any]:
    """Return a JSON-serializable audit; external targets are never opened."""
    path = Path(path)
    if expected_mathtype is not None and expected_mathtype < 0:
        raise ValueError("expected_mathtype must be nonnegative")
    olefile_module = _load_olefile()
    report: dict[str, Any] = {
        "schema_version": 1,
        "file_name": path.name,
        "structure_pass": False,
        "scope_notice": SCOPE_NOTICE,
        "olefile_available": olefile_module is not None,
        "expected_mathtype": expected_mathtype,
        "counts": {"ole_objects": 0, "mathtype_objects": 0, "other_ole_objects": 0},
        "layout_risks": {
            "manual_line_breaks": 0,
            "tabs": 0,
            "text_box_containers": 0,
            "positioned_paragraph_frames": 0,
            "floating_drawings": 0,
            "exact_line_spacing_paragraphs": 0,
            "fixed_height_table_rows": 0,
            "hidden_text_runs_direct": 0,
            "hidden_text_style_definitions": 0,
            "exact_line_spacing_style_definitions": 0,
        },
        "objects": [],
        "issues": [],
    }

    def issue(severity: str, code: str, message: str, **context: Any) -> None:
        report["issues"].append({"severity": severity, "code": code, "message": message, **context})

    try:
        archive = ZipFile(path)
    except (OSError, BadZipFile) as exc:
        issue("error", "INVALID_DOCX_ARCHIVE", "The input could not be opened as a DOCX ZIP archive.",
              exception_type=type(exc).__name__)
        return report

    with archive:
        names = archive.namelist()
        members = set(names)
        for name, count in Counter(names).items():
            if count > 1:
                issue("error", "DUPLICATE_ZIP_MEMBER", "Duplicate ZIP members make target resolution ambiguous.", part=name)
        for required in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
            if required not in members:
                issue("error", "MISSING_DOCX_PART", "A required standard DOCX part is missing.", part=required)

        roots: dict[str, ET.Element] = {}
        for part in sorted(members):
            if part.endswith(".xml") or part.endswith(".rels"):
                try:
                    roots[part] = _xml(archive.read(part))
                except Exception as exc:
                    issue("error", "INVALID_XML_PART", "A package XML part could not be parsed safely.",
                          part=part, exception_type=type(exc).__name__)

        relations: dict[str, dict[str, ET.Element]] = {}
        for part, root in roots.items():
            if not part.endswith(".rels"):
                continue
            mapping = {}
            for rel in root.findall(f"{{{PKG}}}Relationship"):
                rid = rel.get("Id", "")
                if not rid or rid in mapping:
                    issue("error", "INVALID_RELATIONSHIP_ID", "A relationship ID is empty or duplicated.", part=part, rid=rid)
                mapping[rid] = rel
            relations[part] = mapping

        def check_reference(owner: str, rid: str | None, role: str, obj_index: int) -> dict[str, Any]:
            reference: dict[str, Any] = {"rid": rid, "valid": False}
            context = {"part": owner, "object_index": obj_index, "role": role, "rid": rid}
            if not rid:
                issue("error", "MISSING_REFERENCE_ID", "An OLE object or preview is missing its relationship ID.", **context)
                return reference
            rel = relations.get(_rel_part(owner), {}).get(rid)
            if rel is None:
                issue("error", "MISSING_RELATIONSHIP", "A referenced relationship does not exist in the owning part.", **context)
                return reference
            target = rel.get("Target", "")
            reference["relationship_type"] = rel.get("Type", "")
            if rel.get("TargetMode", "Internal").lower() != "internal":
                issue("error", "EXTERNAL_RELATIONSHIP", "MathType objects and previews must be embedded internal parts.", **context)
                return reference
            try:
                resolved = _target_part(owner, target)
            except ValueError:
                issue("error", "INVALID_INTERNAL_TARGET", "The target is not a valid internal package path.", **context)
                return reference
            reference["target_part"] = resolved
            expected_suffix = "/oleObject" if role == "ole" else "/image"
            if not rel.get("Type", "").endswith(expected_suffix):
                issue("error", "WRONG_RELATIONSHIP_TYPE", "The relationship type does not match the object or preview role.", **context)
                return reference
            if resolved not in members:
                issue("error", "MISSING_TARGET_PART", "A relationship target is absent from the ZIP archive.", target_part=resolved, **context)
                return reference
            try:
                size = archive.getinfo(resolved).file_size
                # Reading the member verifies the ZIP CRC too.
                payload = archive.read(resolved)
            except Exception as exc:
                issue("error", "UNREADABLE_TARGET_PART", "An embedded target could not be read or its ZIP checksum failed.",
                      exception_type=type(exc).__name__, **context)
                return reference
            reference["bytes"] = size
            if not payload:
                issue("error", "EMPTY_TARGET_PART", "An embedded object or preview has no data.", **context)
                return reference
            reference["valid"] = True
            return reference

        for part, root in roots.items():
            if part.endswith(".rels"):
                continue
            parents = {child: parent for parent in root.iter() for child in parent}
            risks = report["layout_risks"]
            if part.startswith("word/"):
                for node in root.iter():
                    if node.tag == f"{{{W}}}br" and node.get(f"{{{W}}}type", "textWrapping") == "textWrapping":
                        risks["manual_line_breaks"] += 1
                    elif node.tag == f"{{{W}}}cr":
                        risks["manual_line_breaks"] += 1
                    elif node.tag == f"{{{W}}}tab":
                        risks["tabs"] += 1
                    elif node.tag == f"{{{W}}}txbxContent":
                        risks["text_box_containers"] += 1
                    elif node.tag == f"{{{V}}}textbox" and node.find(f".//{{{W}}}txbxContent") is None:
                        risks["text_box_containers"] += 1
                    elif node.tag == f"{{{W}}}framePr":
                        risks["positioned_paragraph_frames"] += 1
                    elif node.tag == f"{{{WP}}}anchor":
                        risks["floating_drawings"] += 1
                    elif node.tag == f"{{{W}}}p":
                        spacing = node.find(f"{{{W}}}pPr/{{{W}}}spacing")
                        if spacing is not None and spacing.get(f"{{{W}}}lineRule") == "exact":
                            risks["exact_line_spacing_paragraphs"] += 1
                    elif node.tag == f"{{{W}}}tr":
                        heights = node.findall(f"{{{W}}}trPr/{{{W}}}trHeight")
                        if any(height.get(f"{{{W}}}hRule") == "exact" for height in heights):
                            risks["fixed_height_table_rows"] += 1
                    elif node.tag == f"{{{W}}}r":
                        if any(_enabled(node.find(f"{{{W}}}rPr/{{{W}}}{name}")) for name in ("vanish", "webHidden")):
                            risks["hidden_text_runs_direct"] += 1
                    elif node.tag == f"{{{W}}}style":
                        if any(_enabled(node.find(f"{{{W}}}rPr/{{{W}}}{name}")) for name in ("vanish", "webHidden")):
                            risks["hidden_text_style_definitions"] += 1
                        spacing = node.find(f"{{{W}}}pPr/{{{W}}}spacing")
                        if spacing is not None and spacing.get(f"{{{W}}}lineRule") == "exact":
                            risks["exact_line_spacing_style_definitions"] += 1

            for obj in root.iter(f"{{{O}}}OLEObject"):
                obj_index = len(report["objects"]) + 1
                progid = obj.get("ProgID", "")
                is_mathtype = progid.casefold() == MATH_TYPE_PROGID.casefold()
                record: dict[str, Any] = {
                    "index": obj_index, "part": part, "progid": progid,
                    "is_mathtype": is_mathtype, "embedding_type": obj.get("Type", ""),
                    "shape_id": obj.get("ShapeID"), "previews": [],
                }
                report["objects"].append(record)
                report["counts"]["ole_objects"] += 1
                report["counts"]["mathtype_objects" if is_mathtype else "other_ole_objects"] += 1
                if not is_mathtype:
                    # Excel/PowerPoint package objects may use other formats;
                    # do not apply MathType's CFB and preview contract to them.
                    record["validation"] = "outside_mathtype_audit_scope"
                    continue
                if obj.get("Type", "").lower() != "embed":
                    issue("error", "OLE_NOT_EMBEDDED", "An OLE object is linked or does not declare Type=Embed.", part=part, object_index=obj_index)
                ole_ref = check_reference(part, obj.get(f"{{{R}}}id"), "ole", obj_index)
                record["ole"] = ole_ref
                if ole_ref["valid"]:
                    payload = archive.read(ole_ref["target_part"])
                    record["cfb_magic_valid"] = payload.startswith(CFB_MAGIC)
                    if not record["cfb_magic_valid"]:
                        issue("error", "INVALID_CFB_MAGIC", "The embedded OLE payload lacks the Compound File Binary signature.", part=part, object_index=obj_index)
                    elif is_mathtype:
                        native = _native_stream(payload, olefile_module)
                        record["equation_native"] = native
                        if native["status"] not in {"nonempty", "not_checked_dependency_unavailable"}:
                            issue("error", "INVALID_EQUATION_NATIVE", "Equation Native is missing, empty, or cannot be read as a compound-file stream.",
                                  part=part, object_index=obj_index, native_status=native["status"])

                # A normal Word OLE object and its VML preview share w:object.
                container = parents.get(obj)
                while container is not None and container.tag != f"{{{W}}}object":
                    container = parents.get(container)
                if container is None:
                    issue("error", "MISSING_OBJECT_CONTAINER", "The OLE object has no Word object container to associate a preview.", part=part, object_index=obj_index)
                else:
                    preview_ids: list[str | None] = []
                    preview_container = container
                    shapes = list(container.iter(f"{{{V}}}shape"))
                    if shapes and obj.get("ShapeID"):
                        matched = [shape for shape in shapes if shape.get("id") == obj.get("ShapeID")]
                        if len(matched) != 1:
                            issue("error", "PREVIEW_SHAPE_MISMATCH", "ShapeID does not uniquely identify a VML preview shape in the same Word object.",
                                  part=part, object_index=obj_index)
                        else:
                            preview_container = matched[0]
                            record["preview_association"] = "matching_vml_shape_id"
                    else:
                        record["preview_association"] = "word_object_container"
                    for image in preview_container.iter(f"{{{V}}}imagedata"):
                        preview_ids.append(image.get(f"{{{R}}}id"))
                    for image in preview_container.iter(f"{{{A}}}blip"):
                        preview_ids.append(image.get(f"{{{R}}}embed") or image.get(f"{{{R}}}link"))
                    if not preview_ids:
                        issue("error", "MISSING_PREVIEW", "No related VML or DrawingML preview was found in the Word object container.", part=part, object_index=obj_index)
                    for rid in dict.fromkeys(preview_ids):
                        record["previews"].append(check_reference(part, rid, "preview", obj_index))

    count = report["counts"]["mathtype_objects"]
    if expected_mathtype is not None and count != expected_mathtype:
        issue("error", "MATHTYPE_COUNT_MISMATCH", "The number of Equation.DSMT4 objects differs from the expected count.", expected=expected_mathtype, actual=count)
    if count and olefile_module is None:
        issue("warning", "NATIVE_STREAM_NOT_CHECKED", "Install optional olefile to inspect Equation Native streams; only CFB signatures were checked.")
    if not count and expected_mathtype is None:
        issue("warning", "NO_MATHTYPE_OBJECTS", "No Equation.DSMT4 objects were found. OMML equations and pictures are not counted as MathType OLE objects.")
    risk_messages = {
        "manual_line_breaks": "Manual line breaks may preserve PDF line wrapping and prevent natural paragraph reflow.",
        "tabs": "Tabs can create uneven gaps or unstable alignment after edits.",
        "text_box_containers": "Text boxes require rendered checks for clipping, reading order, and overlap.",
        "positioned_paragraph_frames": "Positioned paragraph frames require rendered checks for reflow overflow and overlap.",
        "floating_drawings": "Floating drawings require rendered checks for anchoring and overlap.",
        "exact_line_spacing_paragraphs": "Exact paragraph line spacing can clip tall equations and superscripts.",
        "fixed_height_table_rows": "Exact table row heights can clip equations or wrapped text.",
        "hidden_text_runs_direct": "Some runs are directly marked hidden; verify their intended visibility.",
        "hidden_text_style_definitions": "Hidden text is defined in styles; inherited visibility requires Word verification.",
        "exact_line_spacing_style_definitions": "Styles define exact line spacing; inherited layout requires Word verification.",
    }
    for key, value in report["layout_risks"].items():
        if value:
            issue("warning", "LAYOUT_RISK", risk_messages[key], risk=key, count=value)
    report["structure_pass"] = not any(item["severity"] == "error" for item in report["issues"])
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path, help="DOCX file to audit without modifying it")
    parser.add_argument("--expected-mathtype", type=int, help="Expected number of Equation.DSMT4 objects")
    parser.add_argument("--output", "--out", type=Path, help="Write JSON to this file instead of standard output")
    args = parser.parse_args(argv)
    if args.expected_mathtype is not None and args.expected_mathtype < 0:
        parser.error("--expected-mathtype must be nonnegative")
    if args.output and args.output.resolve() == args.docx.resolve():
        parser.error("The JSON output path must differ from the input DOCX path")
    report = audit_docx(args.docx, args.expected_mathtype)
    text = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)
    return 0 if report["structure_pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
