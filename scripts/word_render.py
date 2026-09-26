"""Render one DOCX with a private Windows Word instance; never save the DOCX.

Requires Windows, desktop Microsoft Word, and pywin32. Concurrent invocations
fail fast via a named mutex. COM calls are synchronous and have no hard timeout;
see references/qa-and-repair.md for stalled-call handling.
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import warnings

MUTEX_NAME = r"Local\PdfToMathTypeDocx.WordRender"
WD_DO_NOT_SAVE = 0
WD_EXPORT_PDF = 17
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


class RenderError(RuntimeError):
    """Word could not produce or cleanly finish a PDF render."""


def _load_windows_modules():
    if os.name != "nt":
        raise RenderError("This renderer requires Windows and desktop Microsoft Word.")
    try:
        import pythoncom
        import win32api
        import win32com.client
        import win32event
    except ImportError as exc:
        raise RenderError("Install pywin32 in this Python environment: python -m pip install pywin32") from exc
    return pythoncom, win32com.client, win32event, win32api


@contextmanager
def _serial_render(win32event, win32api):
    """Serialize this tool across processes without attaching to existing Word."""
    handle = win32event.CreateMutex(None, False, MUTEX_NAME)
    acquired = False
    try:
        result = win32event.WaitForSingleObject(handle, 0)
        # WAIT_ABANDONED means the preceding owner exited; this caller owns it.
        if result not in (0, 0x80):
            raise RenderError("Another word_render process is active. Run Word renders sequentially.")
        acquired = True
        yield
    finally:
        try:
            if acquired:
                win32event.ReleaseMutex(handle)
        finally:
            win32api.CloseHandle(handle)


def _export(source: Path, staged_pdf: Path, pythoncom, client) -> int:
    word = document = None
    initialized = False
    active_error = None
    saved_options = {}
    try:
        pythoncom.CoInitialize()
        initialized = True
        # DispatchEx creates an owned application; never use GetActiveObject
        # or Dispatch, which could attach to the user's open Word session.
        word = client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.ScreenUpdating = False
        word.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        for name in ("UpdateLinksAtOpen", "UpdateLinksAtPrint", "UpdateFieldsAtPrint"):
            saved_options[name] = getattr(word.Options, name)
            setattr(word.Options, name, False)
        document = word.Documents.Open(
            FileName=str(source), ConfirmConversions=False, ReadOnly=True,
            AddToRecentFiles=False, Visible=False, OpenAndRepair=False,
            NoEncodingDialog=True, PasswordDocument="", PasswordTemplate="",
            WritePasswordDocument="", WritePasswordTemplate="",
        )
        document.Repaginate()
        document.ExportAsFixedFormat(
            OutputFileName=str(staged_pdf), ExportFormat=WD_EXPORT_PDF,
            OpenAfterExport=False, OptimizeFor=0,
        )
        return int(document.ComputeStatistics(2))  # wdStatisticPages
    except BaseException as exc:
        active_error = exc
        raise
    finally:
        failures = []
        if document is not None:
            try:
                document.Close(SaveChanges=WD_DO_NOT_SAVE)
            except Exception as exc:
                failures.append(f"document.Close: {exc}")
        document = None
        if word is not None:
            # Word Options can be persisted across instances. Restore the
            # user's values even if opening/exporting/closing the file failed.
            for name, value in saved_options.items():
                try:
                    setattr(word.Options, name, value)
                except Exception as exc:
                    failures.append(f"restore Options.{name}: {exc}")
            try:
                word.Quit(SaveChanges=WD_DO_NOT_SAVE)
            except Exception as exc:
                failures.append(f"owned Word.Quit: {exc}")
        word = None
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception as exc:
                failures.append(f"CoUninitialize: {exc}")
        if failures:
            message = "Word cleanup did not complete: " + "; ".join(failures)
            if active_error is None:
                raise RenderError(message)
            # Preserve the primary error while exposing cleanup problems.
            warnings.warn(message, RuntimeWarning, stacklevel=2)


def _publish(staged_pdf: Path, output: Path, *, force: bool):
    if force:
        # Same-filesystem staging preserves an existing PDF until success.
        os.replace(staged_pdf, output)
        return
    # Exclusive creation handles a second writer appearing after validation.
    # Existing files are never replaced when --force was not requested.
    created = False
    try:
        with output.open("xb") as destination:
            created = True
            with staged_pdf.open("rb") as source:
                shutil.copyfileobj(source, destination)
    except BaseException:
        if created:
            output.unlink(missing_ok=True)
        raise


def render_docx(source: str | Path, output: str | Path | None = None, *, force=False) -> dict:
    source = Path(source).expanduser().resolve(strict=True)
    if not source.is_file() or source.suffix.lower() != ".docx":
        raise ValueError("Input must be an existing .docx file.")
    output = (Path(output).expanduser() if output is not None
              else source.with_name(source.stem + "_word.pdf")).resolve()
    if output.suffix.lower() != ".pdf" or output == source:
        raise ValueError("Output must be a separate .pdf file.")
    if output.exists() and not force:
        raise FileExistsError(f"Output exists; use --force to replace it: {output}")
    if output.exists() and not output.is_file():
        raise ValueError("Output path must identify a file, not a directory.")

    pythoncom, client, win32event, win32api = _load_windows_modules()
    with _serial_render(win32event, win32api):
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix=".word-render-", dir=output.parent) as stage:
            staged_pdf = Path(stage) / "render.pdf"
            pages = _export(source, staged_pdf, pythoncom, client)
            if not staged_pdf.is_file():
                raise RenderError("Word returned without creating a PDF.")
            with staged_pdf.open("rb") as handle:
                if handle.read(5) != b"%PDF-":
                    raise RenderError("Word output does not have a PDF header.")
            _publish(staged_pdf, output, force=force)
    return {"input": str(source), "output": str(output), "pages": pages}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path, help="Existing DOCX; opened read-only")
    parser.add_argument("--output", type=Path, help="PDF path (default: <stem>_word.pdf)")
    parser.add_argument("--force", action="store_true", help="Replace an existing PDF after a successful render")
    args = parser.parse_args(argv)
    try:
        report = render_docx(args.docx, args.output, force=args.force)
    except Exception as exc:
        print(f"word_render: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(report, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
