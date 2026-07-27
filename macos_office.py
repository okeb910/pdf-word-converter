"""Microsoft Office conversion backends for macOS Apple Events."""

from __future__ import annotations

import os
import subprocess
from pathlib import Path
from typing import Callable, Sequence

from engine_models import EngineState, EngineStatus, PathLike, ProgressCallback


OSASCRIPT_PATH = "/usr/bin/osascript"
AUTOMATION_SETTINGS = "系统设置 → 隐私与安全性 → 自动化"
SOURCE_ALREADY_OPEN_MARKER = "PDFWORDCONVERTER_SOURCE_ALREADY_OPEN"
SOURCE_ALREADY_OPEN_DETAIL = (
    "源文件已在 Microsoft Office 中打开。为保护用户正在编辑的文档，本工具不会保存或关闭它；"
    "请先关闭该源文件后重试。"
)
PROBE_TIMEOUT_SECONDS = 20
CONVERSION_TIMEOUT_SECONDS = 600


WORD_PROBE_SCRIPT = """on run argv
    tell application id "com.microsoft.Word"
        return version
    end tell
end run"""


WORD_TO_PDF_SCRIPT = """on run argv
    set sourcePath to item 1 of argv
    set outputPath to item 2 of argv
    set sourceFile to POSIX file sourcePath
    set outputFile to POSIX file outputPath
    set sourceAlias to sourceFile as alias
    set openedDocument to missing value

    tell application id "com.microsoft.Word"
        set sourceAlreadyOpen to false
        repeat with existingDocument in documents
            try
                set existingFullName to (full name of existingDocument) as text
                if existingFullName is sourcePath then
                    set sourceAlreadyOpen to true
                end if
                set existingAlias to existingFullName as alias
                if existingAlias is sourceAlias then
                    set sourceAlreadyOpen to true
                end if
            end try
        end repeat
        if sourceAlreadyOpen then
            error "PDFWORDCONVERTER_SOURCE_ALREADY_OPEN" number -17001
        end if

        try
            set openedDocument to open sourceFile
            save as openedDocument file name outputFile file format format PDF
            close openedDocument saving no
        on error errorMessage number errorNumber
            if openedDocument is not missing value then
                try
                    close openedDocument saving no
                end try
            end if
            error errorMessage number errorNumber
        end try
    end tell
end run"""


POWERPOINT_PROBE_SCRIPT = """on run argv
    tell application id "com.microsoft.Powerpoint"
        return version
    end tell
end run"""


POWERPOINT_TO_PDF_SCRIPT = """on run argv
    set sourcePath to item 1 of argv
    set outputPath to item 2 of argv
    set sourceFile to POSIX file sourcePath
    set outputFile to POSIX file outputPath
    set sourceAlias to sourceFile as alias
    set openedPresentation to missing value

    tell application id "com.microsoft.Powerpoint"
        set sourceAlreadyOpen to false
        repeat with existingPresentation in presentations
            try
                set existingFullName to (full name of existingPresentation) as text
                if existingFullName is sourcePath then
                    set sourceAlreadyOpen to true
                end if
                set existingAlias to existingFullName as alias
                if existingAlias is sourceAlias then
                    set sourceAlreadyOpen to true
                end if
            end try
        end repeat
        if sourceAlreadyOpen then
            error "PDFWORDCONVERTER_SOURCE_ALREADY_OPEN" number -17001
        end if

        try
            set openedPresentation to open sourceFile
            save openedPresentation in outputFile as save as PDF
            close openedPresentation
        on error errorMessage number errorNumber
            if openedPresentation is not missing value then
                try
                    close openedPresentation
                end try
            end if
            error errorMessage number errorNumber
        end try
    end tell
end run"""


class MacOSOfficeError(RuntimeError):
    """A readable native Office failure with its structured engine status."""

    def __init__(self, display_name: str, status: EngineStatus):
        self.status = status
        if status.state is EngineState.PERMISSION_DENIED:
            summary = f"{display_name} 自动化权限被拒绝"
            guidance = (
                f"请在“{AUTOMATION_SETTINGS}”中允许本工具控制 "
                f"{display_name}，然后重试。"
            )
        elif status.state is EngineState.MISSING:
            summary = f"未检测到 {display_name}"
            guidance = (
                f"请安装 {display_name}，或从 Microsoft 官方下载页安装后重试。"
            )
        elif status.state is EngineState.TIMEOUT:
            summary = f"{display_name} 自动化操作超时"
            guidance = (
                f"请关闭 {display_name} 中占用该文件的窗口，确认应用响应后重试。"
            )
        elif status.state is EngineState.LAUNCH_FAILED and status.detail == SOURCE_ALREADY_OPEN_DETAIL:
            summary = f"{display_name} 未开始转换"
            guidance = status.detail
        elif status.state is EngineState.LAUNCH_FAILED:
            summary = f"{display_name} 自动化启动失败"
            guidance = (
                f"请确认 {display_name} 能正常启动，并检查该文档可正常打开，然后重试。"
            )
        else:
            summary = f"{display_name} 转换失败"
            guidance = "请检查应用状态和输入文件后重试。"

        detail = (
            f"\n\n{status.detail}"
            if status.detail and status.detail != guidance
            else ""
        )
        super().__init__(f"{summary}。{guidance}{detail}")


def _default_word_installed() -> bool:
    return _application_exists("Microsoft Word.app")


def _default_powerpoint_installed() -> bool:
    return _application_exists("Microsoft PowerPoint.app")


def _application_exists(app_name: str) -> bool:
    return any(
        candidate.is_dir()
        for candidate in (
            Path("/Applications") / app_name,
            Path.home() / "Applications" / app_name,
        )
    )


def _text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value)


def _failure_status(result: object) -> EngineStatus:
    stdout = _text(getattr(result, "stdout", ""))
    stderr = _text(getattr(result, "stderr", ""))
    returncode = getattr(result, "returncode", None)
    detail = "\n".join(part.strip() for part in (stdout, stderr) if part.strip())
    if not detail:
        detail = f"osascript 返回代码 {returncode}"

    searchable = f"{returncode}\n{stdout}\n{stderr}".casefold()
    if SOURCE_ALREADY_OPEN_MARKER.casefold() in searchable:
        return EngineStatus(EngineState.LAUNCH_FAILED, SOURCE_ALREADY_OPEN_DETAIL)
    permission_markers = (
        "-1743",
        "not authorized",
        "automation denied",
        "not permitted to send apple events",
    )
    state = (
        EngineState.PERMISSION_DENIED
        if any(marker in searchable for marker in permission_markers)
        else EngineState.LAUNCH_FAILED
    )
    return EngineStatus(state, detail)


class _MacOSOfficeBackend:
    id = ""
    display_name = ""
    directions = frozenset()
    _probe_script = ""
    _conversion_script = ""
    _start_message = ""

    def __init__(
        self,
        installed_checker: Callable[[], bool],
        runner: Callable[..., object] = subprocess.run,
    ) -> None:
        self._installed_checker = installed_checker
        self._runner = runner

    def _is_installed(self) -> bool:
        try:
            return bool(self._installed_checker())
        except Exception:
            return False

    def probe(self, deep: bool = False) -> EngineStatus:
        if not self._is_installed():
            return EngineStatus(EngineState.MISSING, f"未检测到 {self.display_name}")
        if not deep:
            return EngineStatus(
                EngineState.UNVERIFIED,
                f"已检测到 {self.display_name}，尚未验证自动化权限",
            )

        try:
            result = self._run_script(
                self._probe_script,
                (),
                timeout=PROBE_TIMEOUT_SECONDS,
            )
        except subprocess.TimeoutExpired as exc:
            return EngineStatus(EngineState.TIMEOUT, _timeout_detail(exc))
        except (OSError, subprocess.SubprocessError) as exc:
            return EngineStatus(EngineState.LAUNCH_FAILED, str(exc))

        if getattr(result, "returncode", 1) == 0:
            detail = _text(getattr(result, "stdout", "")).strip()
            return EngineStatus(EngineState.AVAILABLE, detail)
        return _failure_status(result)

    def convert(
        self,
        source: PathLike,
        output: PathLike,
        progress: ProgressCallback,
    ) -> None:
        if not self._is_installed():
            raise MacOSOfficeError(
                self.display_name,
                EngineStatus(EngineState.MISSING, f"未检测到 {self.display_name}"),
            )

        source_arg = os.fspath(source)
        output_arg = os.fspath(output)
        progress(self._start_message, 5)
        try:
            result = self._run_script(
                self._conversion_script,
                (source_arg, output_arg),
                timeout=CONVERSION_TIMEOUT_SECONDS,
            )
        except subprocess.TimeoutExpired as exc:
            raise MacOSOfficeError(
                self.display_name,
                EngineStatus(EngineState.TIMEOUT, _timeout_detail(exc)),
            ) from exc
        except (OSError, subprocess.SubprocessError) as exc:
            raise MacOSOfficeError(
                self.display_name,
                EngineStatus(EngineState.LAUNCH_FAILED, str(exc)),
            ) from exc

        if getattr(result, "returncode", 1) != 0:
            raise MacOSOfficeError(self.display_name, _failure_status(result))

        progress("完成", 100)

    def _run_script(
        self,
        script: str,
        arguments: Sequence[str],
        *,
        timeout: int,
    ) -> object:
        command = [OSASCRIPT_PATH, "-e", script, *arguments]
        return self._runner(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            shell=False,
            check=False,
        )


def _timeout_detail(exc: subprocess.TimeoutExpired) -> str:
    detail = "\n".join(
        part.strip()
        for part in (_text(getattr(exc, "stdout", "")), _text(getattr(exc, "stderr", "")))
        if part.strip()
    )
    return detail or f"osascript 在 {exc.timeout} 秒内未完成"


class MacOSWordBackend(_MacOSOfficeBackend):
    id = "word_native"
    display_name = "Microsoft Word"
    directions = frozenset({"word_to_pdf"})
    _probe_script = WORD_PROBE_SCRIPT
    _conversion_script = WORD_TO_PDF_SCRIPT
    _start_message = "正在通过 Microsoft Word 导出 PDF..."

    def __init__(
        self,
        installed_checker: Callable[[], bool] = _default_word_installed,
        runner: Callable[..., object] = subprocess.run,
    ) -> None:
        super().__init__(installed_checker, runner)


class MacOSPowerPointBackend(_MacOSOfficeBackend):
    id = "powerpoint_native"
    display_name = "Microsoft PowerPoint"
    directions = frozenset({"powerpoint_to_pdf"})
    _probe_script = POWERPOINT_PROBE_SCRIPT
    _conversion_script = POWERPOINT_TO_PDF_SCRIPT
    _start_message = "正在通过 Microsoft PowerPoint 导出 PDF..."

    def __init__(
        self,
        installed_checker: Callable[[], bool] = _default_powerpoint_installed,
        runner: Callable[..., object] = subprocess.run,
    ) -> None:
        super().__init__(installed_checker, runner)
