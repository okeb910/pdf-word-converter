"""与界面和具体转换引擎解耦的批量转换逻辑。"""

import os
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, List, Optional, Sequence


@dataclass(frozen=True)
class BatchResult:
    source: Path
    output: Optional[Path]
    status: str
    error: str = ""


def deduplicate_paths(paths: Sequence[Path]) -> List[Path]:
    """按规范化绝对路径去重，同时保留用户的选择顺序。"""
    result = []
    seen = set()
    for path in paths:
        normalized = Path(path).expanduser().absolute()
        key = os.path.normcase(os.path.normpath(str(normalized)))
        if key not in seen:
            seen.add(key)
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
        status_changed(BatchResult(source, output, "running"))

        def item_progress(message: str, pct: int, item_index=index) -> None:
            bounded_pct = max(0, min(100, int(pct)))
            overall = int(((item_index + bounded_pct / 100) / total) * 100)
            progress(item_index + 1, total, message, overall)

        try:
            if not source.is_file():
                raise FileNotFoundError(f"源文件不存在: {source}")
            converter(str(source), str(output), item_progress)
            if not output.is_file():
                raise RuntimeError("转换程序未生成输出文件")
            result = BatchResult(source, output, "success")
            progress(index + 1, total, "完成", int(((index + 1) / total) * 100))
        except Exception as exc:
            if output.exists():
                try:
                    output.unlink()
                except OSError:
                    pass
            result = BatchResult(source, output, "failed", str(exc))

        results.append(result)
        status_changed(result)

    return results
