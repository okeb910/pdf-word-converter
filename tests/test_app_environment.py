import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import app_environment as environment
from platform_services import DarwinPlatformServices, WindowsPlatformServices


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

    def test_darwin_validation_does_not_require_windows_or_drag_drop_modules(self):
        imported = []

        def importer(name):
            imported.append(name)
            if name in {"comtypes", "tkinterdnd2"}:
                raise ImportError("platform-specific module must not be required")
            module = Mock()
            if name == "pymupdf":
                module.open = Mock()
            return module

        services = DarwinPlatformServices(platform="darwin", env={})
        self.assertEqual(
            environment.validate_bundled_modules(importer, services=services),
            [],
        )
        self.assertNotIn("comtypes", imported)
        self.assertNotIn("tkinterdnd2", imported)

    def test_windows_validation_requires_comtypes_but_not_drag_drop_module(self):
        imported = []

        def importer(name):
            imported.append(name)
            if name in {"comtypes", "tkinterdnd2"}:
                raise ImportError(f"missing {name}")
            module = Mock()
            if name == "pymupdf":
                module.open = Mock()
            return module

        services = WindowsPlatformServices(platform="win32", env={})
        self.assertEqual(
            environment.validate_bundled_modules(importer, services=services),
            ["comtypes: missing comtypes"],
        )
        self.assertIn("comtypes", imported)
        self.assertNotIn("tkinterdnd2", imported)

    def test_macos_module_failure_message_points_to_source_repair(self):
        services = DarwinPlatformServices(platform="darwin", env={})
        message = environment.describe_module_failures(
            ["PyMuPDF: broken"], services
        )

        self.assertIn("启动工具.command", message)
        self.assertIn("requirements-macos.txt", message)
        self.assertNotIn("便携版", message)

    def test_configure_logging_uses_injected_platform_log_directory(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            log_dir = Path(temp_dir) / "custom-platform-logs"
            services = Mock()
            services.ensure_log_dir.return_value = log_dir
            handler = Mock()

            with patch.object(environment.logging, "FileHandler", return_value=handler) as file_handler, patch.object(
                environment.logging, "basicConfig"
            ) as basic_config, patch.object(environment.logging, "info"):
                log_path = environment.configure_logging(services)

            services.ensure_log_dir.assert_called_once_with()
            self.assertEqual(log_path.parent, log_dir)
            file_handler.assert_called_once_with(log_path, encoding="utf-8")
            self.assertEqual(basic_config.call_args.kwargs["handlers"], [handler])

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
