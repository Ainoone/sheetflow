"""工作表到 Word 文档的通用处理管道。

封装"读取命名区域 → 查找模板 → 自动检测数据类型 → 渲染 Word"的完整流程。
列式数据只忽略所有字段均为空的占位行；含可选空字段的有效记录仍会渲染。
"""

import logging
from collections.abc import Callable
from pathlib import Path
from typing import Any

from .common import pop_required
from .named_ranges import NamedRangeDict
from .records import RecordSet
from .render import render_docx
from .search import require_template
from .user_errors import (
    ERR_MIXED_RECORD_SHAPE,
    ERR_NO_NAMED_RANGES,
    ascii_safe,
    format_user_error,
)

logger = logging.getLogger(__name__)


def _guard_named_ranges(sheet: Any) -> None:
    """若工作表无命名区域，尝试在工作薄中创建错误提示表。

    xlwings "Run main" 模式下，用户可能在错误的工作表上触发脚本。
    此守卫检测空命名区域，在工作薄无密码保护时新增可见的提示表。
    """
    if sheet.names:
        return

    wb = sheet.book
    try:
        tip = wb.sheets.add(name='运行提示')
        tip.range('A1').value = '错误提示'
        tip.range('A2').value = f'当前工作表「{sheet.name}」没有命名区域。'
        tip.range('A3').value = '请切换到包含命名区域的工作表后重新运行。'
        logger.info(f'已在工作薄中创建「运行提示」表')
    except Exception:
        logger.debug('无法创建提示表（工作薄可能有密码保护）')

    raise RuntimeError(
        format_user_error(
            ERR_NO_NAMED_RANGES,
            f"No named ranges on active sheet; sheet={ascii_safe(sheet.name)}",
            f"工作表「{sheet.name}」没有命名区域。请切换到包含命名区域的工作表后重新运行。",
        )
    )


def process_sheet(
    sheet: Any,
    output_dir: Path | Callable[[], Path],
    *,
    filename: Callable[[dict[str, Any]], str],
) -> None:
    """读取工作表命名范围，自动检测列式/标量数据，渲染 Word 文档。

    当前工作表必须包含 Template 命名区域。Template 指定 Word 模板名称；
    其余命名区域作为模板上下文。若所有上下文值都是列表，则按列式数据转为多条记录。

    列式数据（所有值为列表）→ RecordSet 转置为行式记录后渲染；只跳过所有字段
    都是 None 的空白占位行，避免因可选字段为空而静默漏掉有效记录。
    标量数据（单条记录）→ 直接传入 render_docx,由 always_iterable 自动包装。
    列表字段与标量字段不可混用，否则会产生 ERR_MIXED_RECORD_SHAPE。

    Args:
        sheet: xlwings 工作表对象
        output_dir: 输出目录路径，或无参数的延迟创建函数。传入函数时，会在命名区域、
            数据形状、模板和原始输出文件名全部校验通过后才调用，避免失败运行留下
            空目录。
        filename: 从记录生成文件名的可调用对象，如 lambda m: f"{m['CN']}.docx"

    Raises:
        RuntimeError: 当前工作表没有命名区域(ERR_NO_NAMED_RANGES)、
            缺少 Template(ERR_MISSING_NAMED_RANGE)，或混用列表/标量字段
            (ERR_MIXED_RECORD_SHAPE)。
        FileNotFoundError: Template 指定的 Word 模板不存在(ERR_TEMPLATE_NOT_FOUND)。
        KeyError: filename 所需字段或模板渲染字段缺失。
    """
    _guard_named_ranges(sheet)
    named_data = NamedRangeDict(sheet).data
    template_name = pop_required(named_data, 'Template', source=sheet.name)

    list_fields = [key for key, value in named_data.items() if isinstance(value, list)]
    scalar_fields = [key for key, value in named_data.items() if not isinstance(value, list)]
    if list_fields and scalar_fields:
        raise RuntimeError(
            format_user_error(
                ERR_MIXED_RECORD_SHAPE,
                "Named range data mixes list and scalar fields; "
                f"sheet={ascii_safe(sheet.name)}; "
                f"list_fields={', '.join(ascii_safe(key) for key in list_fields)}; "
                f"scalar_fields={', '.join(ascii_safe(key) for key in scalar_fields)}",
                f"工作表「{sheet.name}」命名区域数据混合了列表字段和标量字段。",
            )
        )

    template_path = require_template(template_name)

    if list_fields:
        # 尾部占位行通常整行为空；部分为空则可能只是可选字段缺失，必须保留。
        records = (
            record
            for record in RecordSet(named_data)
            if any(value is not None for value in record.values())
        )
    else:
        records = named_data

    # 将延迟目录工厂继续传给渲染层，使 filename 也能在创建目录前完成预检。
    render_docx(records, template_path, output_dir, filename=filename)
