import subprocess
import unittest
from unittest.mock import Mock

import app_environment as environment


class EnvironmentTests(unittest.TestCase):
    def test_frozen_runtime_root_uses_meipass(self):
        with unittest.mock.patch.object(environment.sys, "frozen", True, create=True), unittest.mock.patch.object(
            environment.sys, "_MEIPASS", r"C:\Temp\bundle", create=True
        ):
            self.assertEqual(str(environment.bundled_runtime_root()), r"C:\Temp\bundle")

    def test_module_validation_reports_failure(self):
        def importer(name):
            if name == "docx":
                raise ImportError("broken")
            module = Mock()
            if name == "pymupdf":
                module.open = Mock()
            return module

        self.assertEqual(
            environment.validate_bundled_modules(importer),
            ["python-docx: broken"],
        )

    def test_exact_office_command(self):
        self.assertEqual(
            environment.build_install_command("office"),
            environment.OFFICE_WINGET_COMMAND,
        )

    def test_exact_libreoffice_command(self):
        self.assertEqual(
            environment.build_install_command("libreoffice"),
            environment.LIBREOFFICE_WINGET_COMMAND,
        )

    def test_missing_winget_is_reported(self):
        self.assertIsNone(environment.find_winget(lambda _name: None))

    def test_install_failure_preserves_return_code(self):
        completed = subprocess.CompletedProcess(["winget"], 42, "", "failed")
        result = environment.run_install_command(("winget",), runner=lambda *_args, **_kwargs: completed)
        self.assertFalse(result.succeeded)
        self.assertEqual(result.returncode, 42)
        self.assertIn("failed", result.message)

    def test_install_start_failure_is_safe(self):
        def fail(*_args, **_kwargs):
            raise FileNotFoundError("missing")

        result = environment.run_install_command(("winget",), runner=fail)
        self.assertFalse(result.succeeded)
        self.assertIsNone(result.returncode)

    def test_official_fallback_uses_expected_url(self):
        opened = []
        url = environment.open_official_download("libreoffice", opened.append)
        self.assertEqual(opened, [environment.LIBREOFFICE_DOWNLOAD_URL])
        self.assertEqual(url, environment.LIBREOFFICE_DOWNLOAD_URL)


if __name__ == "__main__":
    unittest.main()
