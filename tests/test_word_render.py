"""Mock Word ownership, output protection, and failure cleanup on any OS."""
from __future__ import annotations

import importlib.util
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "word_render.py"
SPEC = importlib.util.spec_from_file_location("word_render_under_test", SCRIPT)
render = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(render)


class WordRenderTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "sample.docx"
        self.source.write_bytes(b"synthetic DOCX; COM is mocked")
        self.output = self.root / "result.pdf"
        self.doc = Mock()
        self.doc.ComputeStatistics.return_value = 4
        self.doc.ExportAsFixedFormat.side_effect = self.export_pdf
        self.original_options = dict(UpdateLinksAtOpen=True, UpdateLinksAtPrint=False,
                                     UpdateFieldsAtPrint=True)
        self.word = SimpleNamespace(Options=SimpleNamespace(**self.original_options),
                                    Documents=Mock(), Quit=Mock())
        self.word.Documents.Open.return_value = self.doc
        self.com = Mock()
        self.client = Mock()
        self.client.DispatchEx.return_value = self.word
        self.events = Mock()
        self.events.WaitForSingleObject.return_value = 0
        self.api = Mock()
        self.backend = patch.object(render, "_load_windows_modules", return_value=(
            self.com, self.client, self.events, self.api))
        self.loader = self.backend.start()
        self.addCleanup(self.backend.stop)

    def export_pdf(self, **kwargs):
        Path(kwargs["OutputFileName"]).write_bytes(b"%PDF-1.7\nmock\n%%EOF\n")

    def assert_cleaned(self):
        self.doc.Close.assert_called_once_with(SaveChanges=0)
        self.word.Quit.assert_called_once_with(SaveChanges=0)
        self.com.CoUninitialize.assert_called_once_with()
        self.events.ReleaseMutex.assert_called_once()
        self.api.CloseHandle.assert_called_once()
        self.assertFalse(list(self.root.glob(".word-render-*")))
        for name, value in self.original_options.items():
            self.assertEqual(getattr(self.word.Options, name), value)

    def test_private_invisible_word_disables_macros_and_link_updates_before_open(self):
        def check_open(**kwargs):
            self.assertFalse(self.word.Visible)
            self.assertEqual(self.word.DisplayAlerts, 0)
            self.assertEqual(self.word.AutomationSecurity, 3)
            self.assertFalse(self.word.Options.UpdateLinksAtOpen)
            self.assertFalse(self.word.Options.UpdateLinksAtPrint)
            self.assertFalse(self.word.Options.UpdateFieldsAtPrint)
            self.assertTrue(kwargs["ReadOnly"])
            self.assertFalse(kwargs["AddToRecentFiles"])
            self.assertFalse(kwargs["Visible"])
            return self.doc
        self.word.Documents.Open.side_effect = check_open
        report = render.render_docx(self.source, self.output)
        self.assertEqual(report["pages"], 4)
        self.assertTrue(self.output.read_bytes().startswith(b"%PDF-"))
        self.assertEqual(self.source.read_bytes(), b"synthetic DOCX; COM is mocked")
        self.client.DispatchEx.assert_called_once_with("Word.Application")
        self.client.Dispatch.assert_not_called()
        self.doc.Repaginate.assert_called_once_with()
        self.doc.Save.assert_not_called()
        self.doc.SaveAs2.assert_not_called()
        self.assert_cleaned()

    def test_existing_output_rejected_before_starting_word(self):
        self.output.write_bytes(b"keep me")
        with self.assertRaises(FileExistsError):
            render.render_docx(self.source, self.output)
        self.loader.assert_not_called()
        self.assertEqual(self.output.read_bytes(), b"keep me")

    def test_force_replaces_only_after_successful_export(self):
        self.output.write_bytes(b"old PDF")
        def export(**kwargs):
            self.assertEqual(self.output.read_bytes(), b"old PDF")
            self.export_pdf(**kwargs)
        self.doc.ExportAsFixedFormat.side_effect = export
        render.render_docx(self.source, self.output, force=True)
        self.assertTrue(self.output.read_bytes().startswith(b"%PDF-"))
        self.assert_cleaned()

    def test_export_failure_keeps_existing_output_and_cleans_owned_word(self):
        self.output.write_bytes(b"old PDF")
        self.doc.ExportAsFixedFormat.side_effect = RuntimeError("export failed")
        with self.assertRaisesRegex(RuntimeError, "export failed"):
            render.render_docx(self.source, self.output, force=True)
        self.assertEqual(self.output.read_bytes(), b"old PDF")
        self.assert_cleaned()

    def test_open_failure_still_quits_owned_application(self):
        self.word.Documents.Open.side_effect = RuntimeError("cannot open")
        with self.assertRaisesRegex(RuntimeError, "cannot open"):
            render.render_docx(self.source, self.output)
        self.doc.Close.assert_not_called()
        self.word.Quit.assert_called_once_with(SaveChanges=0)
        self.com.CoUninitialize.assert_called_once_with()
        self.events.ReleaseMutex.assert_called_once()
        self.assertFalse(self.output.exists())
        for name, value in self.original_options.items():
            self.assertEqual(getattr(self.word.Options, name), value)

    def test_options_restored_before_owned_word_quits(self):
        def check_quit(**kwargs):
            for name, value in self.original_options.items():
                self.assertEqual(getattr(self.word.Options, name), value)
        self.word.Quit.side_effect = check_quit
        render.render_docx(self.source, self.output)
        self.assert_cleaned()

    def test_close_failure_does_not_skip_quit_or_publish_output(self):
        self.doc.Close.side_effect = RuntimeError("close failed")
        with self.assertRaisesRegex(render.RenderError, "close failed"):
            render.render_docx(self.source, self.output)
        self.assert_cleaned()
        self.assertFalse(self.output.exists())

    def test_primary_error_preserved_when_cleanup_also_fails(self):
        self.doc.ExportAsFixedFormat.side_effect = RuntimeError("primary failure")
        self.word.Quit.side_effect = RuntimeError("quit failed")
        with self.assertWarnsRegex(RuntimeWarning, "quit failed"):
            with self.assertRaisesRegex(RuntimeError, "primary failure"):
                render.render_docx(self.source, self.output)
        self.assert_cleaned()
        self.assertFalse(self.output.exists())

    def test_busy_mutex_never_launches_word_or_releases_someone_elses_lock(self):
        self.events.WaitForSingleObject.return_value = 258  # WAIT_TIMEOUT
        with self.assertRaisesRegex(render.RenderError, "sequentially"):
            render.render_docx(self.source, self.output)
        self.client.DispatchEx.assert_not_called()
        self.com.CoInitialize.assert_not_called()
        self.events.ReleaseMutex.assert_not_called()
        self.api.CloseHandle.assert_called_once()

    def test_invalid_export_does_not_replace_existing_output(self):
        self.output.write_bytes(b"old PDF")
        def export(**kwargs):
            Path(kwargs["OutputFileName"]).write_bytes(b"not a PDF")
        self.doc.ExportAsFixedFormat.side_effect = export
        with self.assertRaisesRegex(render.RenderError, "PDF header"):
            render.render_docx(self.source, self.output, force=True)
        self.assertEqual(self.output.read_bytes(), b"old PDF")
        self.assert_cleaned()

    def test_output_created_during_render_is_not_overwritten_without_force(self):
        def export(**kwargs):
            self.output.write_bytes(b"another writer")
            self.export_pdf(**kwargs)
        self.doc.ExportAsFixedFormat.side_effect = export
        with self.assertRaises(FileExistsError):
            render.render_docx(self.source, self.output)
        self.assertEqual(self.output.read_bytes(), b"another writer")
        self.assert_cleaned()


if __name__ == "__main__":
    unittest.main()
