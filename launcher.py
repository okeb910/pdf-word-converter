"""Frozen application bootstrap with readable startup failure reporting."""

from __future__ import annotations

import logging
import sys
import traceback

from app_environment import configure_logging, describe_module_failures, validate_bundled_modules
from platform_services import create_platform_services


def show_fatal_error(message: str) -> None:
    """Compatibility wrapper used by older integrations and tests."""
    try:
        create_platform_services().show_fatal_error(message)
    except Exception:
        print(message, file=sys.stderr)


def run() -> int:
    try:
        services = create_platform_services()
    except Exception as exc:
        print(f"不支持当前操作系统: {exc}", file=sys.stderr)
        return 1

    try:
        log_path = configure_logging(services)
    except Exception as exc:
        services.show_fatal_error(
            "无法创建启动日志，程序不能安全启动。\n\n"
            f"日志目录: {services.log_dir}\n错误: {exc}"
        )
        return 1

    failures = validate_bundled_modules(services=services)
    if failures:
        message = describe_module_failures(failures, services)
        logging.error(message)
        services.show_fatal_error(f"{message}\n\n启动日志：{log_path}")
        return 1

    try:
        from pdf_word_converter import main

        main()
        return 0
    except Exception:
        details = traceback.format_exc()
        if services.platform == "darwin":
            recovery = (
                "请重新运行“启动工具.command”修复源码环境，"
                "并查看启动日志。"
            )
        else:
            recovery = "请重新下载安装完整版本。"
        logging.critical("Unhandled startup error\n%s", details)
        services.show_fatal_error(
            f"程序启动时发生错误。{recovery}\n\n"
            f"启动日志：{log_path}\n\n{details[-1500:]}"
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(run())
