"""Operating-system services used by the portable application shell.

This module intentionally performs only path-based application discovery.
Engine-specific modules may do a deeper launch probe later, without making
the shared startup path import ``winreg`` or ``comtypes``.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import webbrowser
from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Mapping, Optional, Sequence, Tuple, Union


APP_NAME = "PDFWordConverter"
OFFICE_DOWNLOAD_URL = "https://www.microsoft.com/microsoft-365/download-office"
LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"

COMMON_REQUIRED_MODULES: Tuple[Tuple[str, str], ...] = (
    ("tkinter", "Tkinter"),
    ("pymupdf", "PyMuPDF"),
    ("docx", "python-docx"),
    ("pptx", "python-pptx"),
)

OFFICE_WINGET_COMMAND = (
    "winget",
    "install",
    "--id",
    "Microsoft.Office",
    "--exact",
    "--source",
    "winget",
    "--accept-package-agreements",
    "--accept-source-agreements",
)
LIBREOFFICE_WINGET_COMMAND = (
    "winget",
    "install",
    "--id",
    "TheDocumentFoundation.LibreOffice",
    "--exact",
    "--source",
    "winget",
    "--accept-package-agreements",
    "--accept-source-agreements",
    "--silent",
)


class InstallerActionKind(str, Enum):
    COMMAND = "command"
    OPEN_URL = "open_url"


@dataclass(frozen=True)
class InstallerAction:
    """A user-approved action for obtaining an external engine."""

    product: str
    kind: InstallerActionKind
    command: Tuple[str, ...] = ()
    url: str = ""

    def __post_init__(self) -> None:
        if not isinstance(self.kind, InstallerActionKind):
            object.__setattr__(self, "kind", InstallerActionKind(self.kind))
        if self.kind is InstallerActionKind.COMMAND:
            if not self.command or self.url:
                raise ValueError("A command installer action needs only a command")
        elif not self.url or self.command:
            raise ValueError("A URL installer action needs only a URL")


@dataclass(frozen=True)
class InstallerActionResult:
    succeeded: bool
    action: InstallerAction
    returncode: Optional[int]
    detail: str = ""


Runner = Callable[..., Any]
Opener = Callable[[str], Any]
Which = Callable[..., Optional[str]]
RequiredModules = Tuple[Tuple[str, str], ...]


def _normalized_platform(value: str) -> str:
    normalized = value.strip().lower()
    if normalized.startswith("win"):
        return "win32"
    if normalized in {"darwin", "mac", "macos"}:
        return "darwin"
    return normalized


def _normalized_product(value: str) -> str:
    normalized = value.strip().lower().replace(" ", "_").replace("-", "_")
    aliases = {
        "microsoft_office": "office",
        "microsoft_word": "word",
        "microsoft_powerpoint": "powerpoint",
        "power_point": "powerpoint",
        "lo": "libreoffice",
        "libre_office": "libreoffice",
    }
    return aliases.get(normalized, normalized)


class PlatformServices(ABC):
    """Small, injectable boundary around OS-specific behavior."""

    def __init__(
        self,
        *,
        platform: Optional[str] = None,
        env: Optional[Mapping[str, str]] = None,
        home: Optional[Union[str, Path]] = None,
        runner: Optional[Runner] = None,
        opener: Optional[Opener] = None,
        which: Optional[Which] = None,
    ) -> None:
        self.platform = _normalized_platform(platform or sys.platform)
        self.env = dict(os.environ if env is None else env)
        self.home = Path.home() if home is None else Path(home)
        self.runner = runner or subprocess.run
        self.opener = opener or self._default_opener()
        self.which = which or shutil.which
        self._env_was_injected = env is not None

    def _default_opener(self) -> Opener:
        return webbrowser.open

    def _find_on_path(self, *names: str) -> Optional[Path]:
        path_value = self.env.get("PATH")
        for name in names:
            try:
                found = self.which(name, path=path_value)
            except TypeError:
                found = self.which(name)
            if found:
                return Path(found)
        return None

    @property
    @abstractmethod
    def required_modules(self) -> RequiredModules:
        """Modules that must be bundled for this platform."""

    @property
    @abstractmethod
    def log_dir(self) -> Path:
        """Directory used for startup and engine-probe logs."""

    def ensure_log_dir(self) -> Path:
        directory = self.log_dir
        directory.mkdir(parents=True, exist_ok=True)
        return directory

    @abstractmethod
    def open_directory(self, path: Union[str, Path]) -> None:
        """Open an existing directory in the native file manager."""

    @abstractmethod
    def show_fatal_error(self, message: str, title: str = "PDF Word PPT Converter 启动失败") -> None:
        """Display a readable startup error without relying on the main GUI."""

    @abstractmethod
    def find_libreoffice(self) -> Optional[Path]:
        """Return a shallow-discovered LibreOffice executable path."""

    @abstractmethod
    def native_app_installed(self, application: str) -> bool:
        """Check native application paths without launching the application."""

    @abstractmethod
    def installer_action(self, product: str) -> InstallerAction:
        """Describe an install or official-download action for user approval."""

    def perform_installer_action(
        self,
        action_or_product: Union[InstallerAction, str],
        *,
        timeout: int = 1800,
    ) -> InstallerActionResult:
        action = (
            action_or_product
            if isinstance(action_or_product, InstallerAction)
            else self.installer_action(action_or_product)
        )

        if action.kind is InstallerActionKind.OPEN_URL:
            try:
                opened = self.opener(action.url)
                succeeded = opened is not False
                detail = "已打开官方下载页" if succeeded else "无法打开官方下载页"
                return InstallerActionResult(succeeded, action, None, detail)
            except Exception as exc:
                return InstallerActionResult(False, action, None, str(exc))

        try:
            completed = self.runner(
                list(action.command),
                capture_output=True,
                text=True,
                timeout=timeout,
                shell=False,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            return InstallerActionResult(False, action, None, str(exc))

        stdout = getattr(completed, "stdout", "") or ""
        stderr = getattr(completed, "stderr", "") or ""
        detail = "\n".join(part.strip() for part in (stdout, stderr) if part.strip())
        returncode = int(getattr(completed, "returncode", 1))
        return InstallerActionResult(
            returncode == 0,
            action,
            returncode,
            detail or ("安装已完成" if returncode == 0 else f"安装程序返回代码 {returncode}"),
        )

    def install_dependency(self, product: str, *, timeout: int = 1800) -> InstallerActionResult:
        """Compatibility-friendly shorthand used by the UI after confirmation."""

        return self.perform_installer_action(product, timeout=timeout)


class WindowsPlatformServices(PlatformServices):
    @property
    def required_modules(self) -> RequiredModules:
        return COMMON_REQUIRED_MODULES + (("comtypes", "comtypes"),)

    @property
    def log_dir(self) -> Path:
        local_app_data = self.env.get("LOCALAPPDATA")
        base = Path(local_app_data) if local_app_data else self.home / "AppData" / "Local"
        return base / APP_NAME / "logs"

    def _default_opener(self) -> Opener:
        startfile = getattr(os, "startfile", None)
        return startfile if startfile is not None else webbrowser.open

    def open_directory(self, path: Union[str, Path]) -> None:
        directory = Path(path).expanduser()
        if not directory.is_dir():
            raise FileNotFoundError(f"目录不存在: {directory}")
        self.opener(str(directory))

    def show_fatal_error(self, message: str, title: str = "PDF Word PPT Converter 启动失败") -> None:
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
        except Exception:
            print(f"{title}: {message}", file=sys.stderr)

    def _program_files_roots(self) -> Tuple[Path, ...]:
        values = []
        for key in ("ProgramFiles", "PROGRAMFILES", "ProgramFiles(x86)", "PROGRAMFILES(X86)"):
            value = self.env.get(key)
            if value and value not in values:
                values.append(value)
        if not values and not self._env_was_injected:
            values.extend((r"C:\Program Files", r"C:\Program Files (x86)"))
        return tuple(Path(value) for value in values)

    def find_libreoffice(self) -> Optional[Path]:
        found = self._find_on_path("libreoffice", "soffice")
        if found:
            return found
        for root in self._program_files_roots():
            candidate = root / "LibreOffice" / "program" / "soffice.exe"
            if candidate.is_file():
                return candidate
        return None

    def _office_executable_candidates(self, executable: str) -> Sequence[Path]:
        candidates = []
        for root in self._program_files_roots():
            office_root = root / "Microsoft Office"
            for version in ("Office16", "Office15", "Office14"):
                candidates.append(office_root / "root" / version / executable)
                candidates.append(office_root / version / executable)
        return candidates

    def native_app_installed(self, application: str) -> bool:
        application = _normalized_product(application)
        if application == "libreoffice":
            return self.find_libreoffice() is not None
        if application == "word":
            return any(path.is_file() for path in self._office_executable_candidates("WINWORD.EXE"))
        if application == "powerpoint":
            return any(path.is_file() for path in self._office_executable_candidates("POWERPNT.EXE"))
        if application == "office":
            return self.native_app_installed("word") or self.native_app_installed("powerpoint")
        raise ValueError(f"Unsupported native application: {application}")

    def installer_action(self, product: str) -> InstallerAction:
        product = _normalized_product(product)
        if product in {"word", "powerpoint"}:
            product = "office"
        commands = {
            "office": OFFICE_WINGET_COMMAND,
            "libreoffice": LIBREOFFICE_WINGET_COMMAND,
        }
        urls = {
            "office": OFFICE_DOWNLOAD_URL,
            "libreoffice": LIBREOFFICE_DOWNLOAD_URL,
        }
        if product not in commands:
            raise ValueError(f"Unsupported installer product: {product}")

        winget = self._find_on_path("winget")
        if winget:
            command = (str(winget), *commands[product][1:])
            return InstallerAction(product, InstallerActionKind.COMMAND, command=command)
        return InstallerAction(product, InstallerActionKind.OPEN_URL, url=urls[product])


class DarwinPlatformServices(PlatformServices):
    @property
    def required_modules(self) -> RequiredModules:
        return COMMON_REQUIRED_MODULES

    @property
    def log_dir(self) -> Path:
        return self.home / "Library" / "Logs" / APP_NAME

    def open_directory(self, path: Union[str, Path]) -> None:
        directory = Path(path).expanduser()
        if not directory.is_dir():
            raise FileNotFoundError(f"目录不存在: {directory}")
        self.runner(["open", str(directory)], check=True, shell=False)

    @staticmethod
    def _apple_script_string(value: str) -> str:
        escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\r", "")
        return escaped.replace("\n", "\\n")

    def show_fatal_error(self, message: str, title: str = "PDF Word PPT Converter 启动失败") -> None:
        script = (
            f'display alert "{self._apple_script_string(title)}" '
            f'message "{self._apple_script_string(message)}" as critical'
        )
        try:
            self.runner(["osascript", "-e", script], check=False, shell=False)
        except Exception:
            print(f"{title}: {message}", file=sys.stderr)

    def _application_roots(self) -> Tuple[Path, Path]:
        return Path("/Applications"), self.home / "Applications"

    def _bundle_exists(self, bundle_name: str) -> bool:
        return any((root / bundle_name).exists() for root in self._application_roots())

    def find_libreoffice(self) -> Optional[Path]:
        for root in self._application_roots():
            candidate = root / "LibreOffice.app" / "Contents" / "MacOS" / "soffice"
            if candidate.is_file():
                return candidate
        return self._find_on_path("libreoffice", "soffice")

    def native_app_installed(self, application: str) -> bool:
        application = _normalized_product(application)
        if application == "libreoffice":
            return self.find_libreoffice() is not None
        if application == "word":
            return self._bundle_exists("Microsoft Word.app")
        if application == "powerpoint":
            return self._bundle_exists("Microsoft PowerPoint.app")
        if application == "office":
            return self.native_app_installed("word") or self.native_app_installed("powerpoint")
        raise ValueError(f"Unsupported native application: {application}")

    def installer_action(self, product: str) -> InstallerAction:
        product = _normalized_product(product)
        if product in {"word", "powerpoint"}:
            product = "office"
        urls = {
            "office": OFFICE_DOWNLOAD_URL,
            "libreoffice": LIBREOFFICE_DOWNLOAD_URL,
        }
        if product not in urls:
            raise ValueError(f"Unsupported installer product: {product}")
        return InstallerAction(product, InstallerActionKind.OPEN_URL, url=urls[product])


def create_platform_services(
    *,
    platform: Optional[str] = None,
    env: Optional[Mapping[str, str]] = None,
    home: Optional[Union[str, Path]] = None,
    runner: Optional[Runner] = None,
    opener: Optional[Opener] = None,
    which: Optional[Which] = None,
) -> PlatformServices:
    platform_name = _normalized_platform(platform or sys.platform)
    service_type = {
        "win32": WindowsPlatformServices,
        "darwin": DarwinPlatformServices,
    }.get(platform_name)
    if service_type is None:
        raise NotImplementedError(f"Unsupported platform: {platform_name}")
    return service_type(
        platform=platform_name,
        env=env,
        home=home,
        runner=runner,
        opener=opener,
        which=which,
    )


get_platform_services = create_platform_services
