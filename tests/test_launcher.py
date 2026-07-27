import unittest
from unittest.mock import Mock, patch

import launcher


class LauncherTests(unittest.TestCase):
    def test_logging_initialization_failure_uses_platform_fatal_error(self):
        services = Mock()
        services.log_dir = "/unwritable/logs"

        with patch.object(launcher, "create_platform_services", return_value=services), patch.object(
            launcher, "configure_logging", side_effect=OSError("permission denied")
        ):
            result = launcher.run()

        self.assertEqual(result, 1)
        services.show_fatal_error.assert_called_once()
        message = services.show_fatal_error.call_args.args[0]
        self.assertIn("/unwritable/logs", message)
        self.assertIn("permission denied", message)

    def test_macos_module_failure_uses_source_repair_guidance(self):
        services = Mock()
        services.platform = "darwin"
        services.log_dir = "/Users/tester/Library/Logs/PDFWordConverter"

        with patch.object(launcher, "create_platform_services", return_value=services), patch.object(
            launcher, "configure_logging", return_value="/tmp/startup.log"
        ), patch.object(
            launcher, "validate_bundled_modules", return_value=["PyMuPDF: broken"]
        ):
            result = launcher.run()

        self.assertEqual(result, 1)
        services.show_fatal_error.assert_called_once()
        message = services.show_fatal_error.call_args.args[0]
        self.assertIn("启动工具.command", message)
        self.assertIn("requirements-macos.txt", message)
        self.assertNotIn("便携版", message)


if __name__ == "__main__":
    unittest.main()
