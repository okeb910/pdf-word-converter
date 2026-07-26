"""Runtime checks and optional external-engine installation helpers."""

from __future__ import annotations

import importlib
import logging
import os
import shutil
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Optional, Sequence


APP_NAME = "PDFWordConverter"
APP_VERSION = "0.4.1"
OFFICE_DOWNLOAD_URL = "https://www.microsoft.com/microsoft-365/download-office"
LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download/download-libreoffice/"

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


@dataclass(frozen=True)
class InstallResult:
    succeeded: bool
    command: tuple[str, ...]
    returncode: Optional[int]
    message: str


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def bundled_runtime_root() -> Path:
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def local_app_data() -> Path:
    value = os.environ.get("LOCALAPPDATA")
    if value:
        return Path(value)
    return Path.home() / "AppData" / "Local"


def configure_logging() -> Path:
    log_dir = local_app_data() / APP_NAME / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"startup-{datetime.now():%Y%m%d}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8")],
        force=True,
    )
    logging.info(
        "%s v%s starting; frozen=%s executable=%s runtime_root=%s",
        APP_NAME,
        APP_VERSION,
        is_frozen(),
        sys.executable,
        bundled_runtime_root(),
    )
    return log_path


def validate_bundled_modules(
    importer: Callable[[str], object] = importlib.import_module,
) -> list[str]:
    """Return user-facing names of missing or damaged embedded modules."""
    required = (
        ("tkinter", "Tkinter"),
        ("pymupdf", "PyMuPDF"),
        ("docx", "python-docx"),
        ("pptx", "python-pptx"),
        ("comtypes", "comtypes"),
        ("tkinterdnd2", "tkinterdnd2"),
    )
    failures = []
    for module_name, display_name in required:
        try:
            module = importer(module_name)
            if module_name == "pymupdf" and not hasattr(module, "open"):
                raise ImportError("missing open()")
        except Exception as exc:
            logging.exception("Bundled module check failed: %s", module_name)
            failures.append(f"{display_name}: {exc}")
    return failures


def find_winget(which: Callable[[str], Optional[str]] = shutil.which) -> Optional[str]:
    return which("winget")


def build_install_command(product: str, winget_path: str = "winget") -> tuple[str, ...]:
    if product == "office":
        command = OFFICE_WINGET_COMMAND
    elif product == "libreoffice":
        command = LIBREOFFICE_WINGET_COMMAND
    else:
        raise ValueError(f"Unsupported product: {product}")
    return (winget_path, *command[1:])


def run_install_command(
    command: Sequence[str],
    runner: Callable[..., subprocess.CompletedProcess] = subprocess.run,
) -> InstallResult:
    command_tuple = tuple(str(part) for part in command)
    logging.info("Starting installer: %s", command_tuple)
    try:
        completed = runner(
            list(command_tuple),
            capture_output=True,
            text=True,
            timeout=1800,
            shell=False,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        logging.exception("Installer could not be started")
        return InstallResult(False, command_tuple, None, str(exc))

    stdout = completed.stdout or ""
    stderr = completed.stderr or ""
    details = "\n".join(part.strip() for part in (stdout, stderr) if part.strip())
    if completed.returncode == 0:
        logging.info("Installer completed successfully: %s", details)
        return InstallResult(True, command_tuple, 0, details or "安装已完成")

    message = details or f"winget 返回代码 {completed.returncode}"
    logging.error("Installer failed (%s): %s", completed.returncode, message)
    return InstallResult(False, command_tuple, completed.returncode, message)


def open_official_download(product: str, opener: Callable[[str], object] = webbrowser.open) -> str:
    if product == "office":
        url = OFFICE_DOWNLOAD_URL
    elif product == "libreoffice":
        url = LIBREOFFICE_DOWNLOAD_URL
    else:
        raise ValueError(f"Unsupported product: {product}")
    logging.info("Opening official download page: %s", url)
    opener(url)
    return url


def describe_module_failures(failures: Iterable[str]) -> str:
    lines = list(failures)
    return (
        "程序内置组件缺失或损坏，无法启动。\n\n"
        + "\n".join(f"- {line}" for line in lines)
        + "\n\n请重新下载安装完整的便携版，或重新运行安装程序。"
    )
