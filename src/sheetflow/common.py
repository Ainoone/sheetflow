"""入口脚本共用的小型流程工具。

本模块集中处理 Run main 错误边界、必需命名区域读取和输出目录创建。
错误文本工具由 user_errors 模块统一提供，并在此兼容转出。
"""

from collections.abc import Callable, Iterable, MutableMapping
from functools import wraps
from pathlib import Path
from typing import Any, ParamSpec, TypeVar

from .smart_path_manager import SmartPathManager
# 保留既有 ``sheetflow.common`` 导入路径，实际实现只由 user_errors 维护。
from .user_errors import (
    ERROR_PREFIX,
    ERR_INVALID_PATH,
    ERR_MISSING_NAMED_RANGE,
    ERR_OUTPUT_DIR,
    ERR_RUN_MAIN_FAILED,
    ascii_safe,
    format_user_error,
)

P = ParamSpec("P")
R = TypeVar("R")


def _missing_message(keys: Iterable[str], source: str | None = None) -> str:
    """生成缺少命名区域时的统一错误信息。"""
    key_text = "、".join(keys)
    if source:
        return f"工作表「{source}」缺少 {key_text} 命名区域。"
    return f"缺少 {key_text} 命名区域。"


def _is_formatted_error(exc: Exception) -> bool:
    """判断异常文本是否已经带有统一错误码前缀。"""
    return str(exc).startswith(ERROR_PREFIX)


def run_main(label: str) -> Callable[[Callable[P, R]], Callable[P, R]]:
    """为 xlwings Run main 入口添加统一错误边界。

    未格式化异常会被包装为 ERR_RUN_MAIN_FAILED RuntimeError；已带 ERR_ 前缀
    的异常保持原样，避免重复包裹。
    """
    def decorator(func: Callable[P, R]) -> Callable[P, R]:
        """为目标 Run main 函数附加统一异常边界。"""
        @wraps(func)
        def wrapper(*args: P.args, **kwargs: P.kwargs) -> R:
            """执行入口函数并按统一格式转换未处理异常。"""
            try:
                return func(*args, **kwargs)
            except Exception as exc:
                if _is_formatted_error(exc):
                    raise
                message = format_user_error(
                    ERR_RUN_MAIN_FAILED,
                    f"Run main failed: {label}; error={ascii_safe(exc)}",
                    f"{label} 运行失败：{exc}",
                )
                raise RuntimeError(message) from None

        return wrapper

    return decorator


def pop_required(
    data: MutableMapping[str, Any],
    key: str,
    *,
    source: str | None = None,
) -> Any:
    """从命名区域字典弹出必需字段，缺失时抛出面向用户的 RuntimeError。

    Args:
        data: 命名区域数据字典。
        key: 必需字段名称。
        source: 可选的工作表名称，用于错误信息。

    Returns:
        被弹出的字段值。

    Raises:
        RuntimeError: data 缺少 key。
    """
    try:
        return data.pop(key)
    except KeyError:
        keys = [key]
        english = f"Missing named range(s): {ascii_safe(key)}"
        if source:
            english += f"; sheet={ascii_safe(source)}"
        raise RuntimeError(
            format_user_error(
                ERR_MISSING_NAMED_RANGE,
                english,
                _missing_message(keys, source),
            )
        ) from None


def pop_required_many(
    data: MutableMapping[str, Any],
    keys: Iterable[str],
    *,
    source: str | None = None,
) -> dict[str, Any]:
    """从命名区域字典一次性弹出多个必需字段。

    若任一字段缺失，函数会先报错，不会修改 data,避免部分 pop 后留下半处理状态。

    Args:
        data: 命名区域数据字典。
        keys: 必需字段名称序列。
        source: 可选的工作表名称，用于错误信息。

    Returns:
        按 keys 顺序弹出的 {字段名: 字段值} 字典。

    Raises:
        RuntimeError: data 缺少任一必需字段。
    """
    keys = tuple(keys)
    missing = [key for key in keys if key not in data]
    if missing:
        english = "Missing named range(s): " + ", ".join(ascii_safe(key) for key in missing)
        if source:
            english += f"; sheet={ascii_safe(source)}"
        raise RuntimeError(
            format_user_error(
                ERR_MISSING_NAMED_RANGE,
                english,
                _missing_message(missing, source),
            )
        )

    return {key: data.pop(key) for key in keys}


def create_output_dir_for_workbook(workbook: Any) -> Path:
    """根据工作簿所在目录创建本次运行的输出目录。

    Args:
        workbook: xlwings 工作簿对象，fullname 需为已保存工作簿的绝对路径。

    Returns:
        SmartPathManager 创建的输出目录路径。

    Raises:
        RuntimeError: 无法创建输出目录。
        ValueError: 工作簿所在目录无效。
    """
    fullname = workbook.fullname
    if not isinstance(fullname, (str, Path)) or not Path(fullname).is_absolute():
        raise ValueError(
            format_user_error(
                ERR_INVALID_PATH,
                f"Workbook must have an absolute saved path; fullname={ascii_safe(fullname)}",
                "请先保存工作簿，再创建输出目录。",
            )
        )
    workbook_dir = Path(fullname).parent
    try:
        output_dir = SmartPathManager(workbook_dir).mkdir_if_needed()
    except OSError as exc:
        raise RuntimeError(
            format_user_error(
                ERR_OUTPUT_DIR,
                f"Unable to create output directory; workbook_dir={ascii_safe(workbook_dir)}; "
                f"error={ascii_safe(exc)}",
                f"无法创建输出目录：{exc}",
            )
        ) from exc
    if output_dir is None:
        raise RuntimeError(
            format_user_error(
                ERR_OUTPUT_DIR,
                f"Unable to create output directory; workbook_dir={ascii_safe(workbook_dir)}",
                "无法创建输出目录，请检查当前目录或桌面是否可写。",
            )
        )
    return output_dir
