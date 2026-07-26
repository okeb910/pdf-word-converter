"""Frozen application bootstrap with readable startup failure reporting."""

from __future__ import annotations

import ctypes
import logging
import sys
import traceback

from app_environment import configure_logging, describe_module_failures, validate_bundled_modules


def show_fatal_error(message: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(None, message, "PDF Word Converter 启动失败", 0x10)
    except Exception:
        print(message, file=sys.stderr)


def run() -> int:
    log_path = configure_logging()
    failures = validate_bundled_modules()
    if failures:
        message = describe_module_failures(failures)
        logging.error(message)
        show_fatal_error(f"{message}\n\n启动日志：{log_path}")
        return 1

    try:
        from pdf_word_converter import main

        main()
        return 0
    except Exception:
        details = traceback.format_exc()
        logging.critical("Unhandled startup error\n%s", details)
        show_fatal_error(
            "程序启动时发生错误。请重新下载安装完整版本。\n\n"
            f"启动日志：{log_path}\n\n{details[-1500:]}"
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(run())
