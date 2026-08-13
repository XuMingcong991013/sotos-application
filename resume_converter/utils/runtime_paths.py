"""
统一解析源码运行和PyInstaller打包后的配置、资源路径。

`.env`始终从程序入口所在目录读取；模板等只读资源在打包后从
PyInstaller临时资源目录读取，避免依赖用户启动程序时的工作目录。
"""

from __future__ import annotations

import sys
from pathlib import Path

def application_directory() -> Path:
    """返回用户可见的程序目录，用于查找外置配置。"""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def resource_path(*parts: str) -> Path:
    """返回源码目录或PyInstaller资源目录中的只读文件路径。"""

    bundle_root = getattr(sys, "_MEIPASS", None)
    root = Path(bundle_root) if bundle_root else application_directory()
    return root.joinpath(*parts)


def application_environment_path() -> Path:
    """返回源码目录或测试版打包资源中的`.env`路径。"""

    return resource_path(".env")
