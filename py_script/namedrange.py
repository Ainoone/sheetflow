"""批量创建工作表级命名区域。

对应 namedrange.xlsx 的 Run main 入口。

工作薄结构：
  - 活动工作表「named range」:命名区域字典含 Named / Range / Sheet_Name 三个键
  - 目标工作表：由 Sheet_Name 指定，命名区域将创建在此表上

本入口只读取活动工作表自身的 sheet.names,避免工作簿级全局名称混入配置。
Sheet_Name 可填写包含空格、括号等标点的实际工作表名，底层创建命名区域时
会按 Excel 引用语法自动转义工作表名前缀。
"""

import xlwings as xw

from sheetflow import (
    NamedRangeDict,
    ascii_safe,
    create_named_ranges_from_dict,
    format_user_error,
    pop_required,
    pop_required_many,
    run_main,
)
from sheetflow.user_errors import ERR_NAMED_RANGE_PARTIAL


def _raise_if_failed(result: dict[str, list[str]]) -> None:
    """部分命名区域创建失败时阻止工作簿保存。"""
    if not result['failed']:
        return

    failed = ", ".join(ascii_safe(name) for name in result['failed'])
    raise RuntimeError(
        format_user_error(
            ERR_NAMED_RANGE_PARTIAL,
            f"Named range creation failed: {failed}",
            f"命名区域创建失败：{failed}",
        )
    )


@run_main("namedrange")
def main() -> None:
    """xlwings 入口：读取当前表配置，批量创建目标表的工作表级命名区域。

    当前活动工作表应提供 Named、Range 和 Sheet_Name 三组命名区域。
    Sheet_Name 指向目标工作表; Named / Range 会被转换为名称到单元格地址的映射。
    目标工作表名可包含括号、空格等需要 Excel 引用转义的字符。

    Raises:
        RuntimeError: 当前配置表缺少 Sheet_Name、Named 或 Range 命名区域
            (ERR_MISSING_NAMED_RANGE)。
            任一命名区域创建失败(ERR_NAMED_RANGE_PARTIAL)时不会保存工作簿。
        ValueError: Named 与 Range 列表长度不一致(ERR_NAMED_RANGE_MAP)。
    """
    wb = xw.Book.caller()
    nrd = NamedRangeDict(wb.sheets.active).data

    source = wb.sheets.active.name
    sheet_name = pop_required(nrd, 'Sheet_Name', source=source)
    named_range_config = pop_required_many(nrd, ('Named', 'Range'), source=source)
    target_sheet = wb.sheets[sheet_name]

    result = create_named_ranges_from_dict(target_sheet, named_range_config)

    print(f"命名区域创建完成：成功 {len(result['ok'])} 个，失败 {len(result['failed'])} 个")
    if result['failed']:
        print(f"失败项: {result['failed']}")
    # 仅在所有命名区域创建成功后保存，避免持久化部分完成的工作簿。
    _raise_if_failed(result)

    wb.save()
    print('工作薄已保存。')
