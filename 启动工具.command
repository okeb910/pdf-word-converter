#!/bin/bash

set -u
set -o pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd -P)"
VENV_DIR="$PROJECT_DIR/.venv-macos"
VENV_PYTHON="$VENV_DIR/bin/python"
REQUIREMENTS_FILE="$PROJECT_DIR/requirements-macos.txt"
STAMP_FILE="$VENV_DIR/.requirements-macos.sha256"
PYTHON_DOWNLOAD_URL="https://www.python.org/downloads/release/python-31210/"

pause_before_exit() {
    if [ -t 0 ]; then
        printf '\n按 Return 键关闭此窗口...'
        IFS= read -r _unused
    fi
}

fail() {
    printf '\n[错误] %s\n' "$1" >&2
    pause_before_exit
    exit 1
}

is_python_312() {
    "$1" -c 'import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 12) else 1)' \
        >/dev/null 2>&1
}

runtime_modules_available() {
    "$1" -c \
        'import tkinter, pymupdf, docx, pptx, PIL.Image, lxml.etree; assert callable(pymupdf.open); assert callable(docx.Document); assert callable(pptx.Presentation)' \
        >/dev/null 2>&1
}

find_python_312() {
    local candidate
    local path_candidate
    local candidates=(
        "/Library/Frameworks/Python.framework/Versions/3.12/bin/python3.12"
        "/opt/homebrew/bin/python3.12"
        "/usr/local/bin/python3.12"
    )

    if path_candidate="$(command -v python3.12 2>/dev/null)"; then
        candidates+=("$path_candidate")
    fi
    if path_candidate="$(command -v python3 2>/dev/null)"; then
        candidates+=("$path_candidate")
    fi

    for candidate in "${candidates[@]}"; do
        if [ -x "$candidate" ] && is_python_312 "$candidate"; then
            printf '%s\n' "$candidate"
            return 0
        fi
    done
    return 1
}

if [ ! -f "$REQUIREMENTS_FILE" ]; then
    fail "缺少 requirements-macos.txt，请重新下载完整源码。"
fi

PYTHON_BIN="$(find_python_312 || true)"
if [ -z "$PYTHON_BIN" ]; then
    printf '[错误] 未找到 Python 3.12。此源码预览不支持其他 Python 版本。\n' >&2
    printf '正在打开 Python 官方 macOS 下载页：%s\n' "$PYTHON_DOWNLOAD_URL" >&2
    open "$PYTHON_DOWNLOAD_URL" >/dev/null 2>&1 || true
    pause_before_exit
    exit 1
fi

if ! "$PYTHON_BIN" -c 'import tkinter' >/dev/null 2>&1; then
    printf '[错误] 当前 Python 3.12 不包含 Tkinter，请安装 python.org 官方 macOS 版本。\n' >&2
    printf '正在打开 Python 官方 macOS 下载页：%s\n' "$PYTHON_DOWNLOAD_URL" >&2
    open "$PYTHON_DOWNLOAD_URL" >/dev/null 2>&1 || true
    pause_before_exit
    exit 1
fi

if [ ! -x "$VENV_PYTHON" ] || ! is_python_312 "$VENV_PYTHON"; then
    printf '[1/3] 正在创建项目内 Python 3.12 环境：%s\n' "$VENV_DIR"
    if [ -d "$VENV_DIR" ]; then
        "$PYTHON_BIN" -m venv --clear "$VENV_DIR" || fail "无法重建 .venv-macos。"
    else
        "$PYTHON_BIN" -m venv "$VENV_DIR" || fail "无法创建 .venv-macos。"
    fi
fi

command -v shasum >/dev/null 2>&1 || fail "系统缺少 shasum，无法校验依赖清单。"
CURRENT_HASH="$(shasum -a 256 "$REQUIREMENTS_FILE" | awk '{print $1}')" \
    || fail "无法计算 requirements-macos.txt 的 SHA-256。"
INSTALLED_HASH=""
if [ -f "$STAMP_FILE" ]; then
    IFS= read -r INSTALLED_HASH < "$STAMP_FILE" || INSTALLED_HASH=""
fi

NEEDS_INSTALL=0
REPAIR_INSTALL=0
if [ "$CURRENT_HASH" != "$INSTALLED_HASH" ]; then
    NEEDS_INSTALL=1
elif ! runtime_modules_available "$VENV_PYTHON"; then
    NEEDS_INSTALL=1
    REPAIR_INSTALL=1
fi

if [ "$NEEDS_INSTALL" -eq 1 ]; then
    if [ "$REPAIR_INSTALL" -eq 1 ]; then
        printf '[2/3] 检测到项目环境损坏，正在重新安装固定版本依赖...\n'
    else
        printf '[2/3] 首次运行或依赖清单已更新，正在安装固定版本依赖...\n'
    fi
    if [ "$REPAIR_INSTALL" -eq 1 ]; then
        "$VENV_PYTHON" -m ensurepip --upgrade >/dev/null 2>&1 || true
    fi
    "$VENV_PYTHON" -m pip install --disable-pip-version-check --upgrade pip \
        || fail "pip 更新失败，请检查网络连接后重试。"
    PIP_REPAIR_ARGS=()
    if [ "$REPAIR_INSTALL" -eq 1 ]; then
        PIP_REPAIR_ARGS=(--force-reinstall)
    fi
    "$VENV_PYTHON" -m pip install --disable-pip-version-check "${PIP_REPAIR_ARGS[@]}" --requirement "$REQUIREMENTS_FILE" \
        || fail "依赖安装失败，请检查网络连接和错误信息。"
    runtime_modules_available "$VENV_PYTHON" \
        || fail "依赖安装完成，但核心模块仍无法加载。请重新下载完整源码后重试。"
    printf '%s\n' "$CURRENT_HASH" > "$STAMP_FILE" \
        || fail "无法写入依赖安装标记。"
else
    printf '[2/3] 依赖清单未变化，跳过安装。\n'
fi

printf '[3/3] 正在启动 PDF ↔ Word/PPT 批量转换工具...\n'
cd "$PROJECT_DIR" || fail "无法进入项目目录。"
"$VENV_PYTHON" "$PROJECT_DIR/launcher.py"
APP_STATUS=$?
if [ "$APP_STATUS" -ne 0 ]; then
    fail "程序启动或运行失败，退出代码：$APP_STATUS。"
fi
