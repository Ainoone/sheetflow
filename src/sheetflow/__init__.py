"""Public package interface for sheetflow.

常用、轻量的工具在导入包时直接暴露；依赖 xlwings 调用上下文或 Excel 自动化的
模块通过 __getattr__ 懒加载，减少普通导入时的副作用。
"""

from .records import RecordSet
from .render import render_docx
from .common import (
    create_output_dir_for_workbook,
    pop_required,
    pop_required_many,
    run_main,
)
from .search import require_template, search_template_file, search_template_file_cached
from .smart_path_manager import SmartPathManager
from .time_period import generate_period_range
from .user_errors import ascii_safe, format_user_error

__all__ = [
    "NamedRangeDict",
    "RecordSet",
    "SmartPathManager",
    "ascii_safe",
    "build_named_range_map",
    "create_named_ranges_from_dict",
    "create_output_dir_for_workbook",
    "format_user_error",
    "create_sheet_named_ranges",
    "generate_period_range",
    "generate_personal_income_tax",
    "pop_required",
    "pop_required_many",
    "process_sheet",
    "render_docx",
    "require_template",
    "run_main",
    "search_template_file",
    "search_template_file_cached",
]


def __getattr__(name: str):
    """按需加载较重的公共对象。

    该函数服务于模块级懒加载：调用方仍可从 sheetflow 直接导入 __all__ 中的名称，
    但只有实际访问这些名称时才导入对应子模块。

    Args:
        name: 调用方访问的属性名。

    Returns:
        对应的公共对象。

    Raises:
        AttributeError: name 不是 sheetflow 暴露的公共名称。
    """
    if name in {
        "NamedRangeDict",
        "build_named_range_map",
        "create_named_ranges_from_dict",
        "create_sheet_named_ranges",
    }:
        from .named_ranges import (
            NamedRangeDict,
            build_named_range_map,
            create_named_ranges_from_dict,
            create_sheet_named_ranges,
        )

        exports = {
            "NamedRangeDict": NamedRangeDict,
            "build_named_range_map": build_named_range_map,
            "create_named_ranges_from_dict": create_named_ranges_from_dict,
            "create_sheet_named_ranges": create_sheet_named_ranges,
        }
        return exports[name]

    if name == "generate_personal_income_tax":
        from .personal_income_tax import generate_personal_income_tax

        return generate_personal_income_tax

    if name == "process_sheet":
        from .pipeline import process_sheet

        return process_sheet

    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
