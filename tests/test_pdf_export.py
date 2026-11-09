"""Backend selection and failure isolation checks; no Office installation needed."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch


sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "fastresume_pdf_export_under_test", ROOT / "scripts" / "pdf_export.py"
)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError("Cannot load PDF export backend")
backend = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(backend)


class PdfBackendTests(unittest.TestCase):
    def test_auto_fallback_removes_partial_output_before_next_backend(self) -> None:
        attempts = []

        def word_failure(docx, pdf, temporary, timeout):
            attempts.append("word")
            pdf.write_bytes(b"incomplete Word output")
            raise RuntimeError("Word is unavailable")

        def libreoffice_success(docx, pdf, temporary, timeout, office_path):
            attempts.append("libreoffice")
            self.assertFalse(pdf.exists(), "A partial export must not reach the next backend")
            pdf.write_bytes(b"completed PDF fixture")

        with tempfile.TemporaryDirectory(prefix="fastresume-pdf-fallback-check-") as temporary:
            root = Path(temporary)
            source, output = root / "input.docx", root / "resume.pdf"
            source.write_bytes(b"source fixture; conversion is mocked")
            # Replace this module's os reference, not os.name globally: Path and
            # TemporaryDirectory must retain the actual platform on Linux CI.
            with patch.object(backend, "os", SimpleNamespace(name="nt")), \
                    patch.object(backend, "_convert_word", side_effect=word_failure), \
                    patch.object(backend, "_convert_libreoffice", side_effect=libreoffice_success):
                backend.convert_pdf(source, output, workdir=root)
            self.assertEqual(["word", "libreoffice"], attempts)
            self.assertEqual(b"completed PDF fixture", output.read_bytes())
            self.assertEqual({source, output}, set(root.iterdir()))

    def test_explicit_office_path_selects_libreoffice_in_auto_mode(self) -> None:
        selected_path = "custom-office/soffice"

        def libreoffice_success(docx, pdf, temporary, timeout, office_path):
            self.assertEqual(selected_path, office_path)
            pdf.write_bytes(b"completed PDF fixture")

        with tempfile.TemporaryDirectory(prefix="fastresume-pdf-selection-check-") as temporary:
            root = Path(temporary)
            source, output = root / "input.docx", root / "resume.pdf"
            source.write_bytes(b"source fixture; conversion is mocked")
            with patch.object(backend, "os", SimpleNamespace(name="nt")), \
                    patch.object(backend, "_convert_word") as word, \
                    patch.object(backend, "_convert_libreoffice", side_effect=libreoffice_success) as libreoffice:
                backend.convert_pdf(source, output, office_path=selected_path, workdir=root)
            word.assert_not_called()
            libreoffice.assert_called_once()
            self.assertEqual(b"completed PDF fixture", output.read_bytes())
            self.assertEqual({source, output}, set(root.iterdir()))

    def test_invalid_timeout_starts_no_backend_and_leaves_no_files(self) -> None:
        with tempfile.TemporaryDirectory(prefix="fastresume-pdf-timeout-check-") as temporary:
            root = Path(temporary)
            source, output = root / "input.docx", root / "resume.pdf"
            source.write_bytes(b"source fixture")
            with patch.object(backend, "_convert_word") as word, \
                    patch.object(backend, "_convert_libreoffice") as libreoffice:
                for timeout in (0, -1, float("nan"), float("inf"), True):
                    with self.subTest(timeout=timeout), self.assertRaises(RuntimeError):
                        backend.convert_pdf(source, output, timeout=timeout, workdir=root)
            word.assert_not_called()
            libreoffice.assert_not_called()
            self.assertEqual({source}, set(root.iterdir()))

    def test_all_backend_failures_preserve_source_and_remove_partial_exports(self) -> None:
        original = b"source must remain unchanged"

        def word_failure(docx, pdf, temporary, timeout):
            pdf.write_bytes(b"partial Word output")
            raise RuntimeError("Word refused an existing session")

        def libreoffice_failure(docx, pdf, temporary, timeout, office_path):
            self.assertFalse(pdf.exists())
            pdf.write_bytes(b"partial LibreOffice output")
            raise RuntimeError("LibreOffice conversion failed")

        with tempfile.TemporaryDirectory(prefix="fastresume-pdf-failure-check-") as temporary:
            root = Path(temporary)
            source, output = root / "input.docx", root / "resume.pdf"
            source.write_bytes(original)
            with patch.object(backend, "os", SimpleNamespace(name="nt")), \
                    patch.object(backend, "_convert_word", side_effect=word_failure), \
                    patch.object(backend, "_convert_libreoffice", side_effect=libreoffice_failure):
                with self.assertRaises(RuntimeError) as raised:
                    backend.convert_pdf(source, output, workdir=root)
            self.assertIn("Word refused an existing session", str(raised.exception))
            self.assertIn("LibreOffice conversion failed", str(raised.exception))
            self.assertEqual(original, source.read_bytes())
            self.assertEqual({source}, set(root.iterdir()))


if __name__ == "__main__":
    unittest.main()
