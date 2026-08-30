"""工商登记 Word 文档批量生成。

对应 工商.xlsx 的 Run main 入口。
从活动工作表读取命名区域，先校验配置与模板，再创建目录并生成工商登记类 Word 文档。
"""

import xlwings as xw

from sheetflow import create_output_dir_for_workbook, process_sheet, run_main


@run_main("business_registration")
def main() -> None:
    """xlwings 入口：从当前工作表生成工商登记类 Word 文档。

    当前工作表必须包含 Template 命名区域，以及模板渲染所需的业务字段。
    输出目录由当前工作簿所在目录和 SmartPathManager 的规则共同决定，并在工作表、
    数据形状、模板与原始输出文件名校验通过后才创建。

    Raises:
        RuntimeError: 无法创建输出目录(ERR_OUTPUT_DIR)、当前工作表没有命名区域
            (ERR_NO_NAMED_RANGES)，或命名区域数据形状不一致(ERR_MIXED_RECORD_SHAPE)。
        FileNotFoundError: Template 指定的 Word 模板不存在(ERR_TEMPLATE_NOT_FOUND)。
        KeyError: 文件名生成或模板渲染所需字段缺失。
    """
    wb = xw.Book.caller()
    # 延迟目录创建，避免配置或模板错误留下空的时间戳目录。
    process_sheet(
        wb.sheets.active,
        lambda: create_output_dir_for_workbook(wb),
        filename=lambda m: f"{m['CN']}.docx",
    )
