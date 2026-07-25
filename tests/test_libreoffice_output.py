import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pdf_word_converter as converter


class CompletedProcess:
    returncode = 0
    stderr = ""


class LibreOfficeOutputTests(unittest.TestCase):
    def test_docx_to_pdf_moves_isolated_output_to_requested_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "report.docx"
            existing = folder / "report.pdf"
            requested = folder / "report_converted.pdf"
            source.write_bytes(b"docx")
            existing.write_bytes(b"existing")

            def fake_run(command, **_kwargs):
                output_dir = Path(command[command.index("--outdir") + 1])
                (output_dir / "report.pdf").write_bytes(b"new-pdf")
                return CompletedProcess()

            with patch.object(converter, "_get_lo_path", return_value=r"C:\FakeLO\soffice.exe"), patch.object(
                converter.subprocess, "run", side_effect=fake_run
            ):
                converter.docx_to_pdf_via_libreoffice(
                    str(source), str(requested), lambda *_args: None
                )

            self.assertEqual(existing.read_bytes(), b"existing")
            self.assertEqual(requested.read_bytes(), b"new-pdf")

    def test_pdf_to_word_moves_isolated_output_to_requested_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "report.pdf"
            existing = folder / "report.docx"
            requested = folder / "report_converted_2.docx"
            source.write_bytes(b"pdf")
            existing.write_bytes(b"existing")

            def fake_run(command, **_kwargs):
                output_dir = Path(command[command.index("--outdir") + 1])
                (output_dir / "report.docx").write_bytes(b"new-docx")
                return CompletedProcess()

            with patch.object(converter, "_get_lo_path", return_value=r"C:\FakeLO\soffice.exe"), patch.object(
                converter.subprocess, "run", side_effect=fake_run
            ):
                converter.pdf_to_word_via_libreoffice(
                    str(source), str(requested), lambda *_args: None
                )

            self.assertEqual(existing.read_bytes(), b"existing")
            self.assertEqual(requested.read_bytes(), b"new-docx")


if __name__ == "__main__":
    unittest.main()
