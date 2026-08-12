"""
各处理阶段共用的异常日志追加工具。
"""

from __future__ import annotations

import traceback
from datetime import datetime
from pathlib import Path


def write_error_log(
    output_dir: Path,
    input_file_path: Path,
    stage: str,
) -> Path:
    """把当前异常的完整堆栈追加到按日期命名的日志。"""

    output_dir.mkdir(parents=True, exist_ok=True)
    now = datetime.now()
    error_file = output_dir / f"error_{now:%Y%m%d}.txt"
    error_content = (
        f"{'=' * 80}\n"
        f"时间：{now:%Y-%m-%d %H:%M:%S}\n"
        f"处理阶段：{stage}\n"
        f"输入文件：{input_file_path.resolve()}\n"
        f"异常信息：\n{traceback.format_exc()}\n"
    )

    with error_file.open("a", encoding="utf-8") as file:
        file.write(error_content)

    return error_file
