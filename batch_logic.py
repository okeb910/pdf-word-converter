"""与界面和具体转换引擎解耦的批量转换逻辑。"""

import os
import threading
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, List, Optional, Sequence
from uuid import uuid4


@dataclass(frozen=True)
class BatchResult:
    source: Path
    output: Optional[Path]
    status: str
    error: str = ""


def deduplicate_paths(paths: Sequence[Path]) -> List[Path]:
    """按文件身份或规范化绝对路径去重，同时保留用户的选择顺序。"""
    result = []
    seen_files = set()
    seen_paths = set()
    for path in paths:
        normalized = Path(path).expanduser().absolute()
        try:
            stat_result = normalized.stat()
            inode = getattr(stat_result, "st_ino", 0)
            device = getattr(stat_result, "st_dev", 0)
            file_key = (device, inode) if inode else None
        except OSError:
            file_key = None

        if file_key is not None:
            if file_key in seen_files:
                continue
            seen_files.add(file_key)
        else:
            path_key = unicodedata.normalize(
                "NFC",
                os.path.normcase(os.path.normpath(str(normalized))),
            )
            if path_key in seen_paths:
                continue
            seen_paths.add(path_key)

        result.append(normalized)
    return result


def resolve_output_path(
    input_path: Path,
    target_suffix: str,
    output_dir: Optional[Path] = None,
) -> Path:
    """生成不覆盖现有文件的目标路径。"""
    input_path = Path(input_path)
    target_suffix = target_suffix if target_suffix.startswith(".") else f".{target_suffix}"
    target_suffix = target_suffix.lower()
    target_dir = Path(output_dir) if output_dir is not None else input_path.parent

    primary = target_dir / f"{input_path.stem}{target_suffix}"
    if not primary.exists():
        return primary

    converted = target_dir / f"{input_path.stem}_converted{target_suffix}"
    if not converted.exists():
        return converted

    index = 2
    while True:
        candidate = target_dir / f"{input_path.stem}_converted_{index}{target_suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def _staging_output_path(output_path: Path) -> Path:
    """Return a private same-directory path that still has the target suffix."""

    output = Path(output_path)
    return output.with_name(
        f".{output.stem}.{uuid4().hex}.partial{output.suffix}"
    )


def _cleanup_failed_output(
    output_path: Path,
    attempts: int = 3,
    retry_delay: float = 0.05,
) -> str:
    """Remove a rejected output, or isolate it under an unmistakable name."""

    output = Path(output_path)
    if not output.exists():
        return ""

    last_error = None
    for attempt in range(max(1, attempts)):
        try:
            output.unlink()
            return ""
        except FileNotFoundError:
            return ""
        except OSError as exc:
            last_error = exc
            if attempt + 1 < max(1, attempts):
                time.sleep(max(0.0, retry_delay) * (attempt + 1))

    if not output.exists():
        return ""

    for index in range(1, 1000):
        suffix = ".failed" if index == 1 else f".failed_{index}"
        quarantine = output.with_name(output.name + suffix)
        if quarantine.exists():
            continue
        try:
            output.rename(quarantine)
        except OSError as exc:
            last_error = exc
            break
        return f"不完整输出无法删除，已隔离为: {quarantine}"

    return (
        f"不完整输出清理失败，文件仍位于: {output}"
        + (f" ({last_error})" if last_error else "")
    )


def _publish_staged_output(staging_path: Path, output_path: Path) -> None:
    """Atomically publish a verified file without overwriting an existing output."""

    staging = Path(staging_path)
    output = Path(output_path)
    try:
        os.link(staging, output)
    except FileExistsError as exc:
        raise RuntimeError(f"输出路径在转换期间已被占用: {output}") from exc
    except OSError as exc:
        raise RuntimeError(f"无法安全发布转换结果: {output} ({exc})") from exc

    try:
        staging.unlink()
    except OSError:
        _cleanup_failed_output(staging)


def run_conversion_batch(
    input_paths: Sequence[Path],
    target_suffix: str,
    output_dir: Optional[Path],
    converter: Callable[[str, str, Callable[[str, int], None]], None],
    cancel_event: threading.Event,
    progress: Callable[[int, int, str, int], None],
    status_changed: Callable[[BatchResult], None],
) -> List[BatchResult]:
    """串行执行转换；单项失败不会中断其余任务。"""
    paths = deduplicate_paths(input_paths)
    total = len(paths)
    results = []

    for index, source in enumerate(paths):
        if cancel_event.is_set():
            cancelled = BatchResult(source, None, "cancelled")
            results.append(cancelled)
            status_changed(cancelled)
            continue

        output = resolve_output_path(source, target_suffix, output_dir)
        staging = _staging_output_path(output)
        status_changed(BatchResult(source, output, "running"))

        def item_progress(message: str, pct: int, item_index=index) -> None:
            bounded_pct = max(0, min(100, int(pct)))
            overall = int(((item_index + bounded_pct / 100) / total) * 100)
            progress(item_index + 1, total, message, overall)

        try:
            if not source.is_file():
                raise FileNotFoundError(f"源文件不存在: {source}")
            converter(str(source), str(staging), item_progress)
            if not staging.is_file():
                raise RuntimeError("转换程序未生成输出文件")
            _publish_staged_output(staging, output)
            result = BatchResult(source, output, "success")
            progress(index + 1, total, "完成", int(((index + 1) / total) * 100))
        except Exception as exc:
            cleanup_message = _cleanup_failed_output(staging)
            error = str(exc) + (f"\n{cleanup_message}" if cleanup_message else "")
            # A rejected artifact must never be exposed as a usable output path.
            result = BatchResult(source, None, "failed", error)
        results.append(result)
        status_changed(result)

    return results
