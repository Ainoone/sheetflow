"""税务申报 Word/Excel 文档批量生成。

对应 税务.xlsx 的 Run main 入口。
支持普通税务模板渲染和个税压缩包两种模式；普通 Word 模板会在创建输出目录前校验。
"""

from collections.abc import Callable, Iterable
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw

from sheetflow import (
    NamedRangeDict,
    create_output_dir_for_workbook,
    generate_period_range,
    generate_personal_income_tax,
    pop_required,
    pop_required_many,
    render_docx,
    require_template,
    run_main,
)

from sheetflow.user_errors import ERR_OUTPUT_PATH, format_user_error

PERSONAL_INCOME_TAX_PACKAGE = '个税压缩包'
PERIOD_KEYS = ('start', 'end', 'freq')


@run_main("tax")
def main() -> None:
    """xlwings 入口：按时间区间生成税务 Word 文档或个税压缩包。

    当前工作表必须包含 Template、start、end、freq 命名区域。Template 为
    "个税压缩包" 时直接调用个税 Excel 生成流程；其他值按同名 Word 模板渲染
    周期性税务文档，并在创建输出目录前确认 Word 模板存在。

    Raises:
        RuntimeError: 当前工作表缺少必需命名区域(ERR_MISSING_NAMED_RANGE)，
            或无法创建输出目录(ERR_OUTPUT_DIR)。
        FileNotFoundError: Template 指定的模板文件不存在(ERR_TEMPLATE_NOT_FOUND)。
        TypeError: start 或 end 不是支持的日期类型。
        ValueError: freq 不是支持的时间切分频率、日期元组无效，或 Word 分支的
            CN 不是非空公司名称(ERR_OUTPUT_PATH)。
        KeyError: 文件名生成或模板渲染所需字段缺失。
    """
    wb = xw.Book.caller()

    named_data = NamedRangeDict(wb.sheets.active).data
    sheet_name = wb.sheets.active.name
    template_name = pop_required(named_data, 'Template', source=sheet_name)
    timeseries = pop_required_many(named_data, PERIOD_KEYS, source=sheet_name)
    # 此调用会立即校验日期与频率；返回的一次性迭代器再由下游消费。
    periods = generate_period_range(timeseries)

    # 个税流程生成 Excel 后压缩；普通 Word 流程先确认模板存在，避免空输出目录。
    if template_name == PERSONAL_INCOME_TAX_PACKAGE:
        output_dir = create_output_dir_for_workbook(wb)
        generate_personal_income_tax(named_data, periods, output_dir)
    else:
        template_path = require_template(template_name)
        _render_periodic_tax_docx(
            named_data, periods, template_path,
            lambda: create_output_dir_for_workbook(wb),
        )


def _render_periodic_tax_docx(
    named_data: dict[str, Any],
    periods: Iterable[tuple[date, date]],
    template_path: Path,
    output_dir: Path | Callable[[], Path],
) -> None:
    """按时间区间渲染普通税务 Word 文档。

    Args:
        named_data: 当前工作表命名区域读取出的业务字段。
        periods: 已切分好的时间区间，每项可解包为 (start, end)。
        template_path: 已在创建输出目录前校验通过的 Word 模板路径。
        output_dir: 本次运行的输出目录，或文件名预检通过后调用的目录创建函数。
    """
    company_name = named_data.get('CN')
    if not isinstance(company_name, str) or not company_name.strip():
        raise ValueError(
            format_user_error(
                ERR_OUTPUT_PATH,
                "CN must be a non-empty company name for the output filename",
                "生成文件名前，CN 必须填写非空的公司名称。",
            )
        )

    records = (
        {**named_data, 'start': start, 'end': end}
        for start, end in periods
    )
    render_docx(
        records,
        template_path,
        output_dir,
        filename=lambda m: f"{company_name[:6]}_{m['end'].year}年{m['end'].month:02}月.docx",
    )
