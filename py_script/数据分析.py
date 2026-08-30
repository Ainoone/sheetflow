"""单记录 Excel 模板填充。

对应 数据分析.xlsx 的 Run main 入口。

前提: namedrange.py 已完成命名区域创建，目标工作表已手动追加工作表级
Template 命名区域。
"""

from typing import Any

import xlwings as xw

from sheetflow import (
    NamedRangeDict,
    ascii_safe,
    create_output_dir_for_workbook,
    format_user_error,
    pop_required,
    require_template,
    run_main,
)
from sheetflow.user_errors import ERR_TEMPLATE_WRITE_FAILED


def _write_fields(sheet: Any, data: dict[str, Any]) -> list[tuple[str, Exception]]:
    """写入模板字段并返回失败项，不在此处保存工作簿。"""
    failed = []
    for key, value in data.items():
        try:
            sheet.range(key).value = value
        except Exception as exc:
            failed.append((key, exc))
    return failed


@run_main("data_analysis")
def main() -> None:
    """xlwings 入口：将当前工作表的一条记录填入 Excel 模板。

    若当前活动表是 named range 配置表，会先切换到第一个非配置表。
    当前目标表必须包含 Template 命名区域；可选 Sheet 命名区域用于指定模板中的
    目标工作表，否则默认写入模板首个工作表。

    Raises:
        RuntimeError: 当前工作表缺少 Template 命名区域(ERR_MISSING_NAMED_RANGE)，
            无法创建输出目录(ERR_OUTPUT_DIR)，或模板字段写入失败
            (ERR_TEMPLATE_WRITE_FAILED)。
        FileNotFoundError: Template 指定的 Excel 模板不存在(ERR_TEMPLATE_NOT_FOUND)。
    """
    wb = xw.Book.caller()

    # 若误留在 named range 配置表，自动切换到目标表。
    if wb.sheets.active.name == 'named range':
        target = next((s for s in wb.sheets if s.name != 'named range'), None)
        if target is not None:
            target.activate()

    nrd = NamedRangeDict(wb.sheets.active).data

    template_name = pop_required(nrd, 'Template', source=wb.sheets.active.name)
    template_path = require_template(template_name, suffix='xlsx')
    output_dir = create_output_dir_for_workbook(wb)

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        twb = app.books.open(template_path)
        try:
            # Template 只用于选择模板；Sheet 只用于选择模板工作表，
            # 其余命名区域才作为模板字段写入。
            sheet_name = nrd.pop('Sheet', None)
            sht = twb.sheets[sheet_name] if sheet_name else twb.sheets[0]

            failed = _write_fields(sht, nrd)
            if failed:
                # 任一字段写入失败都不保存正式文件，避免生成字段不完整的结果。
                details = "; ".join(
                    f"{ascii_safe(key)}: {ascii_safe(exc)}"
                    for key, exc in failed
                )
                raise RuntimeError(
                    format_user_error(
                        ERR_TEMPLATE_WRITE_FAILED,
                        f"Template fields failed: {details}",
                        f"模板字段写入失败：{details}",
                    )
                )

            output_path = output_dir / f'DONE_{template_name}.xlsx'
            twb.save(output_path)
        finally:
            twb.close()
