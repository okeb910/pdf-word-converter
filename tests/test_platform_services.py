import subprocess
import tempfile
import unittest
from pathlib import Path

from platform_services import (
    APP_NAME,
    LIBREOFFICE_DOWNLOAD_URL,
    OFFICE_DOWNLOAD_URL,
    DarwinPlatformServices,
    InstallerActionKind,
    WindowsPlatformServices,
    create_platform_services,
)


class RecordingRunner:
    def __init__(self, result=None):
        self.calls = []
        self.result = result or subprocess.CompletedProcess([], 0, "ok", "")

    def __call__(self, *args, **kwargs):
        self.calls.append((args, kwargs))
        return self.result


class PlatformServiceTests(unittest.TestCase):
    def test_factory_selects_windows_and_darwin(self):
        self.assertIsInstance(create_platform_services(platform="Windows", env={}), WindowsPlatformServices)
        self.assertIsInstance(create_platform_services(platform="macOS", env={}), DarwinPlatformServices)
        with self.assertRaises(NotImplementedError):
            create_platform_services(platform="linux", env={})

    def test_required_modules_keep_comtypes_windows_only(self):
        windows = WindowsPlatformServices(platform="win32", env={})
        darwin = DarwinPlatformServices(platform="darwin", env={})
        self.assertIn(("comtypes", "comtypes"), windows.required_modules)
        self.assertNotIn(("comtypes", "comtypes"), darwin.required_modules)
        self.assertNotIn(("tkinterdnd2", "TkinterDnD2"), windows.required_modules)
        self.assertNotIn(("tkinterdnd2", "TkinterDnD2"), darwin.required_modules)

    def test_platform_log_directories(self):
        windows = WindowsPlatformServices(
            platform="win32",
            env={"LOCALAPPDATA": r"C:\Users\tester\AppData\Local"},
            home=Path("unused"),
        )
        darwin = DarwinPlatformServices(platform="darwin", env={}, home=Path("/Users/tester"))
        self.assertEqual(
            windows.log_dir,
            Path(r"C:\Users\tester\AppData\Local") / APP_NAME / "logs",
        )
        self.assertEqual(darwin.log_dir, Path("/Users/tester/Library/Logs") / APP_NAME)

    def test_windows_open_directory_uses_injected_startfile_semantics(self):
        opened = []
        with tempfile.TemporaryDirectory() as temp_dir:
            service = WindowsPlatformServices(platform="win32", env={}, opener=opened.append)
            service.open_directory(temp_dir)
        self.assertEqual(opened, [temp_dir])

    def test_darwin_open_directory_uses_argument_list_without_shell(self):
        runner = RecordingRunner()
        with tempfile.TemporaryDirectory() as temp_dir:
            service = DarwinPlatformServices(platform="darwin", env={}, runner=runner)
            service.open_directory(temp_dir)
        args, kwargs = runner.calls[0]
        self.assertEqual(args[0], ["open", temp_dir])
        self.assertIs(kwargs["shell"], False)

    def test_darwin_libreoffice_checks_user_applications_before_path(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            home = Path(temp_dir)
            soffice = home / "Applications" / "LibreOffice.app" / "Contents" / "MacOS" / "soffice"
            soffice.parent.mkdir(parents=True)
            soffice.touch()
            service = DarwinPlatformServices(
                platform="darwin",
                env={},
                home=home,
                which=lambda *_args, **_kwargs: "/path/should/not/win",
            )
            self.assertEqual(service.find_libreoffice(), soffice)

    def test_darwin_native_office_probe_is_shallow(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            home = Path(temp_dir)
            (home / "Applications" / "Microsoft Word.app").mkdir(parents=True)
            service = DarwinPlatformServices(platform="darwin", env={}, home=home)
            self.assertTrue(service.native_app_installed("word"))
            self.assertTrue(service.native_app_installed("office"))
            self.assertFalse(service.native_app_installed("powerpoint"))

    def test_darwin_installers_only_open_official_pages(self):
        opened = []
        runner = RecordingRunner()
        service = DarwinPlatformServices(
            platform="darwin", env={}, runner=runner, opener=opened.append
        )
        office = service.installer_action("word")
        libreoffice = service.installer_action("libreoffice")
        self.assertEqual((office.kind, office.url), (InstallerActionKind.OPEN_URL, OFFICE_DOWNLOAD_URL))
        self.assertEqual(
            (libreoffice.kind, libreoffice.url),
            (InstallerActionKind.OPEN_URL, LIBREOFFICE_DOWNLOAD_URL),
        )
        result = service.perform_installer_action(libreoffice)
        self.assertTrue(result.succeeded)
        self.assertEqual(opened, [LIBREOFFICE_DOWNLOAD_URL])
        self.assertEqual(runner.calls, [])

    def test_windows_winget_action_preserves_argument_list_and_shell_false(self):
        runner = RecordingRunner()
        service = WindowsPlatformServices(
            platform="win32",
            env={},
            runner=runner,
            which=lambda name, **_kwargs: r"C:\Windows\winget.exe" if name == "winget" else None,
        )
        action = service.installer_action("office")
        self.assertIs(action.kind, InstallerActionKind.COMMAND)
        self.assertEqual(action.command[0], r"C:\Windows\winget.exe")
        self.assertIn("Microsoft.Office", action.command)

        result = service.perform_installer_action(action)
        self.assertTrue(result.succeeded)
        args, kwargs = runner.calls[0]
        self.assertEqual(args[0], list(action.command))
        self.assertIs(kwargs["shell"], False)

    def test_windows_without_winget_falls_back_to_official_page(self):
        service = WindowsPlatformServices(
            platform="win32", env={}, which=lambda *_args, **_kwargs: None
        )
        action = service.installer_action("libreoffice")
        self.assertEqual(action.kind, InstallerActionKind.OPEN_URL)
        self.assertEqual(action.url, LIBREOFFICE_DOWNLOAD_URL)


if __name__ == "__main__":
    unittest.main()
