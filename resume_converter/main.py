"""
简历标准化工具的交互式批量处理入口。

用户选择输入文件夹、输出文件夹和模板模式后，程序依次处理输入文件夹
当前层级中的全部受支持简历，并在终端展示当前文件、阶段和最终汇总。
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from resume_pipeline import ResumeConverter
from utils.restore_markdown import SUPPORTED_SUFFIXES


Converter = Callable[[str, str, str, bool], str | None]
InputReader = Callable[[str], str]
ReadabilityChecker = Callable[[Path], tuple[bool, str]]


@dataclass(frozen=True)
class TemplateOption:
    """命令行菜单中的一个可选模板模式。"""

    label: str
    template_tag: str
    with_photo: bool


@dataclass(frozen=True)
class BatchResult:
    """一次文件夹批处理的结果汇总。"""

    total: int
    succeeded: tuple[Path, ...]
    failed: tuple[Path, ...]
    skipped: tuple[Path, ...]
    stopped_early: bool = False


# 后续增加模板时只需继续注册菜单项，不需要改动交互和批处理流程。
TEMPLATE_OPTIONS: dict[str, TemplateOption] = {
    "1": TemplateOption(
        label="SOTOS - 有照片占位",
        template_tag="SOTOS",
        with_photo=True,
    ),
    "2": TemplateOption(
        label="SOTOS - 无照片",
        template_tag="SOTOS",
        with_photo=False,
    ),
    "3": TemplateOption(
        label="优族 - 无照片",
        template_tag="优族",
        with_photo=False,
    ),
    "4": TemplateOption(
        label="奇瑞",
        template_tag="奇瑞",
        with_photo=False,
    ),
}


def main() -> None:
    """运行交互式命令行批处理。"""

    _print_welcome()

    try:
        input_directory = _prompt_input_directory()
        output_directory = _prompt_output_directory(input_directory)
        template = _prompt_template_option()
        files, ignored = discover_resume_files(input_directory)

        _print_task_confirmation(
            input_directory,
            output_directory,
            template,
            files,
            ignored,
        )

        if not files:
            print("\n没有找到可处理的简历，程序已结束。")
            return

        duplicate_stems = _find_duplicate_stems(files)
        if duplicate_stems:
            _print_duplicate_stem_warning(duplicate_stems)
            return

        result = process_folder(
            files=files,
            output_directory=output_directory,
            template=template,
        )
        _print_batch_summary(result, output_directory)
    except KeyboardInterrupt:
        print("\n\n已收到用户中断，程序安全退出。")


def discover_resume_files(
    input_directory: Path,
) -> tuple[list[Path], list[Path]]:
    """扫描文件夹当前层级，分别返回可处理和被忽略的文件。"""

    supported: list[Path] = []
    ignored: list[Path] = []

    for path in input_directory.iterdir():
        if not path.is_file():
            continue
        # Word/WPS打开DOCX时通常会产生~$开头的临时锁文件。
        if path.name.startswith("~$"):
            ignored.append(path.resolve())
        elif path.suffix.lower() in SUPPORTED_SUFFIXES:
            supported.append(path.resolve())
        else:
            ignored.append(path.resolve())

    sort_key = lambda path: path.name.casefold()
    return sorted(supported, key=sort_key), sorted(ignored, key=sort_key)


def _find_duplicate_stems(files: list[Path]) -> dict[str, list[Path]]:
    """查找会生成同名过程文件和最终Word的原文件。"""

    grouped: dict[str, list[Path]] = {}
    for path in files:
        grouped.setdefault(path.stem.casefold(), []).append(path)
    return {
        stem: paths
        for stem, paths in grouped.items()
        if len(paths) > 1
    }


def process_folder(
    files: list[Path],
    output_directory: Path,
    template: TemplateOption,
    converter: Converter = ResumeConverter,
    input_reader: InputReader = input,
    readability_checker: ReadabilityChecker | None = None,
) -> BatchResult:
    """依次处理已发现的简历，并保证单份失败不影响后续文件。"""

    succeeded: list[Path] = []
    failed: list[Path] = []
    skipped: list[Path] = []
    total = len(files)
    check_readability = readability_checker or _check_file_readability

    print("\n" + "=" * 72)
    print(f"开始批量处理，共 {total} 份简历")
    print("=" * 72)

    for index, source_path in enumerate(files, start=1):
        print("\n" + "-" * 72)
        print(f"正在处理 [{index}/{total}]：{source_path.name}")
        print(f"原始文件：{source_path}")
        print("-" * 72)

        access_action = _resolve_file_access(
            source_path,
            check_readability,
            input_reader,
        )
        if access_action == "skip":
            skipped.append(source_path)
            print(f"\n  [跳过] {source_path.name}")
            continue
        if access_action == "stop":
            skipped.extend(files[index - 1 :])
            print("\n  已按用户选择提前结束，未处理剩余文件。")
            return BatchResult(
                total=total,
                succeeded=tuple(succeeded),
                failed=tuple(failed),
                skipped=tuple(skipped),
                stopped_early=True,
            )

        try:
            result = converter(
                str(source_path),
                str(output_directory),
                template.template_tag,
                template.with_photo,
            )
        except Exception as error:
            # 公共入口通常会自行记录错误；这里额外保护批处理不中断。
            print(f"\n  [失败] 出现未处理异常：{error}")
            result = None

        if result:
            succeeded.append(Path(result).resolve())
            print(f"\n  [完成] {source_path.name}")
            print(f"  最终简历：{Path(result).resolve()}")
        else:
            failed.append(source_path)
            print(f"\n  [失败] {source_path.name}")
            print("  已继续处理下一份简历；详细原因请查看过程数据错误日志。")

    return BatchResult(
        total=total,
        succeeded=tuple(succeeded),
        failed=tuple(failed),
        skipped=tuple(skipped),
    )


def _check_file_readability(path: Path) -> tuple[bool, str]:
    """尝试实际读取一个字节，判断文件是否被其他程序独占。"""

    try:
        with path.open("rb") as file:
            file.read(1)
        return True, ""
    except (PermissionError, OSError) as error:
        return False, str(error)


def _resolve_file_access(
    source_path: Path,
    readability_checker: ReadabilityChecker,
    input_reader: InputReader,
) -> str:
    """文件被占用时让用户重试、跳过或结束批处理。"""

    while True:
        readable, error_message = readability_checker(source_path)
        if readable:
            return "ready"

        print("\n  [文件暂时无法读取]")
        print(f"  文件：{source_path.name}")
        print("  可能原因：文件正被其他程序独占，或当前用户没有读取权限。")
        if error_message:
            print(f"  系统提示：{error_message}")
        print("  请关闭占用该文件的软件后选择重试。")
        print("    R. 重试当前文件")
        print("    S. 跳过当前文件")
        print("    Q. 结束本次批处理")

        choice = input_reader("  请输入 R、S 或 Q：\n> ").strip().lower()
        if choice == "r":
            continue
        if choice == "s":
            return "skip"
        if choice == "q":
            return "stop"
        print("  无效选项，请重新输入。")


def _prompt_input_directory() -> Path:
    """循环读取一个真实存在的输入文件夹。"""

    while True:
        raw_value = input(
            "请输入原始简历所在文件夹路径：\n> "
        ).strip()
        path = _path_from_user_input(raw_value)

        if not raw_value:
            print("路径不能为空，请重新输入。\n")
        elif not path.exists():
            print(f"找不到该路径：{path}\n")
        elif not path.is_dir():
            print(f"该路径不是文件夹：{path}\n")
        else:
            return path.resolve()


def _prompt_output_directory(input_directory: Path) -> Path:
    """循环读取输出文件夹，尚不存在时自动创建。"""

    while True:
        raw_value = input(
            "\n请输入结果输出文件夹路径（不存在时会自动创建）：\n> "
        ).strip()
        path = _path_from_user_input(raw_value)

        if not raw_value:
            print("路径不能为空，请重新输入。")
            continue
        if path.exists() and not path.is_dir():
            print(f"该路径不是文件夹：{path}")
            continue
        if path.resolve() == input_directory.resolve():
            print("输出文件夹不能与原始简历文件夹相同，请重新输入。")
            continue

        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            print(f"无法创建输出文件夹：{error}")
            continue

        return path.resolve()


def _prompt_template_option() -> TemplateOption:
    """展示可扩展模板菜单，并只接受已注册的数字选项。"""

    print("\n请选择标准简历模板：")
    for number, option in TEMPLATE_OPTIONS.items():
        print(f"  {number}. {option.label}")

    while True:
        choice = input("请输入选项数字：\n> ").strip()
        option = TEMPLATE_OPTIONS.get(choice)
        if option:
            return option
        valid_options = "、".join(TEMPLATE_OPTIONS)
        print(f"无效选项，请输入：{valid_options}。\n")


def _path_from_user_input(value: str) -> Path:
    """兼容用户直接粘贴带英文或中文引号的Windows路径。"""

    cleaned = value.strip().strip('"\'“”‘’')
    return Path(cleaned).expanduser()


def _print_welcome() -> None:
    """输出简洁但信息完整的启动说明。"""

    print("=" * 72)
    print("简历标准化批量处理工具")
    print("=" * 72)
    print("本工具会依次完成：")
    print("  1. 百度文档解析与 Markdown 结构恢复")
    print("  2. 简历信息提取和证据校验")
    print("  3. 生成标准 Word 简历及待补充信息 Excel")
    print("\n当前支持：PDF、DOCX、JPG、JPEG、PNG")
    print("说明：本次只扫描所选文件夹当前层级，不进入子文件夹。\n")
    print("文件处于打开状态但允许读取时可正常处理；被独占时会提示重试或跳过。\n")


def _print_task_confirmation(
    input_directory: Path,
    output_directory: Path,
    template: TemplateOption,
    files: list[Path],
    ignored: list[Path],
) -> None:
    """在正式调用外部服务前展示任务范围。"""

    print("\n" + "=" * 72)
    print("任务信息确认")
    print("=" * 72)
    print(f"输入文件夹：{input_directory}")
    print(f"输出文件夹：{output_directory}")
    print(f"模板模式：  {template.label}")
    print(f"待处理文件：{len(files)} 个")

    if ignored:
        print(f"忽略其他文件：{len(ignored)} 个")
        for path in ignored:
            print(f"  - {path.name}")


def _print_batch_summary(
    result: BatchResult,
    output_directory: Path,
) -> None:
    """输出成功、失败和结果位置汇总。"""

    print("\n" + "=" * 72)
    print("批量处理结束")
    print("=" * 72)
    print(f"总计：{result.total} 份")
    print(f"成功：{len(result.succeeded)} 份")
    print(f"失败：{len(result.failed)} 份")
    print(f"跳过：{len(result.skipped)} 份")

    if result.failed:
        print("\n失败文件：")
        for path in result.failed:
            print(f"  - {path.name}")

    if result.skipped:
        print("\n跳过或未处理文件：")
        for path in result.skipped:
            print(f"  - {path.name}")

    if result.stopped_early:
        print("\n本次任务由用户提前结束。")

    print("\n结果位置：")
    print(f"  最终简历：{output_directory / 'final_resumes'}")
    print(
        "  补充清单："
        f"{output_directory / 'final_resumes' / '待补充信息.xlsx'}"
    )
    print(f"  过程数据：{output_directory / 'process_data'}")


def _print_duplicate_stem_warning(
    duplicate_stems: dict[str, list[Path]],
) -> None:
    """在调用API前阻止同stem文件相互覆盖。"""

    print("\n检测到以下文件会生成同名结果，为避免覆盖，本次尚未开始处理：")
    for paths in duplicate_stems.values():
        print(f"  文件名主体：{paths[0].stem}")
        for path in paths:
            print(f"    - {path.name}")
    print("请先重命名这些文件，确保每份简历的文件名主体不同后再运行。")


if __name__ == "__main__":
    main()
