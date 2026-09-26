"""Synthetic DOCX tests: no private documents or reference material required."""
from __future__ import annotations

import importlib.util
from contextlib import redirect_stderr
from io import BytesIO, StringIO
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "audit_docx.py"
SPEC = importlib.util.spec_from_file_location("audit_docx_under_test", SCRIPT)
audit = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(audit)


def make_docx(path: Path, *, count=1, missing_ole=False, missing_preview=False,
              bad_rid=False, external=False, bad_magic=False, layout_risks=False,
              external_preview=False, wrong_shape=False):
    objects = []
    relationships = []
    files = {}
    for index in range(count):
        ole_id, image_id = f"rIdOle{index}", f"rIdImage{index}"
        target = "https://example.invalid/equation.bin" if external else f"embeddings/equation{index}.bin"
        mode = ' TargetMode="External"' if external else ""
        relationships += [
            f'<Relationship Id="{ole_id}" Type="{audit.R}/oleObject" Target="{target}"{mode}/>',
            f'<Relationship Id="{image_id}" Type="{audit.R}/image" Target="media/preview{index}.png"'
            + (' TargetMode="External"' if external_preview else '') + '/>',
        ]
        rid = "rIdMissing" if bad_rid else ole_id
        shape_id = "nonexistentShape" if wrong_shape else f"shape{index}"
        objects.append(
            '<w:r><w:object><v:shape id="shape' + str(index) + '">'
            f'<v:imagedata r:id="{image_id}"/></v:shape>'
            f'<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="{shape_id}" r:id="{rid}"/>'
            '</w:object></w:r>'
        )
        if not missing_ole:
            files[f"word/embeddings/equation{index}.bin"] = (b"not-a-cfb" if bad_magic else audit.CFB_MAGIC + b"\x00" * 504)
        if not missing_preview:
            # A nonempty synthetic image payload is sufficient for relationship
            # testing. Image decoding and visual quality are explicitly outside
            # this structural auditor's contract.
            files[f"word/media/preview{index}.png"] = b"\x89PNG\r\n\x1a\nsynthetic"
    properties = '<w:pPr><w:spacing w:line="200" w:lineRule="exact"/></w:pPr>' if layout_risks else ''
    extra = '<w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t><w:br/><w:tab/></w:r>' if layout_risks else ''
    files["word/document.xml"] = (
        f'<w:document xmlns:w="{audit.W}" xmlns:r="{audit.R}" xmlns:o="{audit.O}" xmlns:v="{audit.V}">'
        '<w:body><w:p>' + properties + ''.join(objects) + extra + '</w:p></w:body></w:document>'
    ).encode()
    files["word/_rels/document.xml.rels"] = (
        f'<Relationships xmlns="{audit.PKG}">' + ''.join(relationships) + '</Relationships>'
    ).encode()
    files["_rels/.rels"] = (
        f'<Relationships xmlns="{audit.PKG}"><Relationship Id="rIdMain" Type="{audit.R}/officeDocument" '
        'Target="word/document.xml"/></Relationships>'
    ).encode()
    files["[Content_Types].xml"] = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    ).encode()
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)


class AuditDocxTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "synthetic.docx"
        self.optional = patch.object(audit, "_load_olefile", return_value=None)
        self.optional.start()
        self.addCleanup(self.optional.stop)

    def codes(self, report):
        return {issue["code"] for issue in report["issues"]}

    def test_count_and_valid_relationships(self):
        make_docx(self.path, count=2)
        report = audit.audit_docx(self.path, expected_mathtype=2)
        self.assertTrue(report["structure_pass"])
        self.assertEqual(report["counts"]["mathtype_objects"], 2)
        self.assertTrue(all(obj["cfb_magic_valid"] for obj in report["objects"]))
        self.assertIn("NATIVE_STREAM_NOT_CHECKED", self.codes(report))
        self.assertIn("does not prove", report["scope_notice"])

    def test_expected_count_mismatch(self):
        make_docx(self.path)
        report = audit.audit_docx(self.path, expected_mathtype=2)
        self.assertFalse(report["structure_pass"])
        self.assertIn("MATHTYPE_COUNT_MISMATCH", self.codes(report))

    def test_missing_relationship(self):
        make_docx(self.path, bad_rid=True)
        report = audit.audit_docx(self.path)
        self.assertFalse(report["structure_pass"])
        self.assertIn("MISSING_RELATIONSHIP", self.codes(report))

    def test_missing_ole_target(self):
        make_docx(self.path, missing_ole=True)
        report = audit.audit_docx(self.path)
        self.assertFalse(report["structure_pass"])
        self.assertIn("MISSING_TARGET_PART", self.codes(report))

    def test_missing_preview_target(self):
        make_docx(self.path, missing_preview=True)
        report = audit.audit_docx(self.path)
        self.assertFalse(report["structure_pass"])
        self.assertEqual(report["objects"][0]["previews"][0]["valid"], False)

    def test_preview_shape_must_match_ole_shape_id(self):
        make_docx(self.path, wrong_shape=True)
        report = audit.audit_docx(self.path)
        self.assertFalse(report["structure_pass"])
        self.assertIn("PREVIEW_SHAPE_MISMATCH", self.codes(report))

    def test_internal_path_resolution_rejects_escape_and_urls(self):
        self.assertEqual(audit._target_part("word/document.xml", "media/preview.png"), "word/media/preview.png")
        for value in ("../../outside.bin", "https://example.invalid/file.bin", "file:///tmp/file.bin"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                audit._target_part("word/document.xml", value)

    def test_external_ole_and_preview_are_rejected(self):
        for kwargs in ({"external": True}, {"external_preview": True}):
            with self.subTest(kwargs=kwargs):
                make_docx(self.path, **kwargs)
                report = audit.audit_docx(self.path)
                self.assertFalse(report["structure_pass"])
                self.assertIn("EXTERNAL_RELATIONSHIP", self.codes(report))

    def test_invalid_compound_file_signature(self):
        make_docx(self.path, bad_magic=True)
        report = audit.audit_docx(self.path)
        self.assertFalse(report["structure_pass"])
        self.assertIn("INVALID_CFB_MAGIC", self.codes(report))

    def test_fixed_spacing_and_hidden_text_are_warnings(self):
        make_docx(self.path, layout_risks=True)
        report = audit.audit_docx(self.path)
        self.assertTrue(report["structure_pass"])
        self.assertEqual(report["layout_risks"]["exact_line_spacing_paragraphs"], 1)
        self.assertEqual(report["layout_risks"]["manual_line_breaks"], 1)
        self.assertEqual(report["layout_risks"]["tabs"], 1)
        self.assertEqual(report["layout_risks"]["hidden_text_runs_direct"], 1)
        self.assertIn("LAYOUT_RISK", self.codes(report))

    def test_optional_native_stream_is_required_to_be_nonempty(self):
        class FakeOle:
            data = b""
            def __init__(self, source): pass
            def exists(self, name): return name == "Equation Native"
            def get_size(self, name): return len(self.data)
            def openstream(self, name): return BytesIO(self.data)
            def close(self): pass
        make_docx(self.path)
        for data, expected in ((b"", False), (b"synthetic native payload", True)):
            FakeOle.data = data
            with patch.object(audit, "_load_olefile", return_value=SimpleNamespace(OleFileIO=FakeOle)):
                report = audit.audit_docx(self.path)
            self.assertEqual(report["structure_pass"], expected)
            self.assertEqual(report["objects"][0]["equation_native"]["status"], "nonempty" if data else "empty")

    def test_cli_json_output(self):
        make_docx(self.path)
        output = Path(self.temp.name) / "audit.json"
        for flag in ("--output", "--out"):
            with self.subTest(flag=flag):
                status = audit.main([str(self.path), "--expected-mathtype", "1", flag, str(output)])
                self.assertEqual(status, 0)
                self.assertEqual(json.loads(output.read_text(encoding="utf-8"))["counts"]["mathtype_objects"], 1)

    def test_invalid_archive_reports_failure(self):
        self.path.write_bytes(b"not a ZIP file")
        report = audit.audit_docx(self.path)
        self.assertFalse(report["structure_pass"])
        self.assertIn("INVALID_DOCX_ARCHIVE", self.codes(report))

    def test_cli_refuses_to_overwrite_input(self):
        make_docx(self.path)
        original = self.path.read_bytes()
        equivalent_path = self.path.parent / "unused" / ".." / self.path.name
        with redirect_stderr(StringIO()), self.assertRaises(SystemExit) as raised:
            audit.main([str(self.path), "--out", str(equivalent_path)])
        self.assertEqual(raised.exception.code, 2)
        self.assertEqual(self.path.read_bytes(), original)

    def test_cli_creates_output_parent_directories(self):
        make_docx(self.path)
        output = Path(self.temp.name) / "new" / "nested" / "audit.json"
        self.assertFalse(output.parent.exists())
        self.assertEqual(audit.main([str(self.path), "--out", str(output)]), 0)
        self.assertTrue(json.loads(output.read_text(encoding="utf-8"))["structure_pass"])


if __name__ == "__main__":
    unittest.main()
