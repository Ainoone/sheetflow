"""个人所得税 Excel 模板批量生成模块。

按调用方提供的时间段填充 Excel 模板并保存为 .xls 文件，支持新旧个税自动切换。

典型用例：
    >>> from datetime import date
    >>> from pathlib import Path
    >>> data = {'CC': '91310115...', 'CN': '上海某公司', 'IDN': '342123...'}
    >>> periods = [(date(2019, 1, 1), date(2019, 1, 31))]
    >>> output_dir = generate_personal_income_tax(data, periods, Path('output'))
"""

from datetime import date
import logging
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw

from .search import search_template_file_cached
from .autozip import auto_zip
from .user_errors import ERR_OUTPUT_PATH, ERR_TEMPLATE_NOT_FOUND, ascii_safe, format_user_error

logger = logging.getLogger(__name__)

TEMPLATE_NAMES = {2018: '2018', 2019: '2019'}


@auto_zip()
def generate_personal_income_tax(
    data: dict[str, Any],
    periods: Iterable[tuple[date, date]],
    output_dir: Path,
) -> Path:
    """按调用方提供的时间段生成个人所得税计算表。

    Args:
        data: 公共数据字典（来自命名区域），如 CC、CN、IDN 等。每个键都会被
            直接传给 `sht.range(key)`，因此必须对应个税 .xls 模板中的命名区域
            或单元格地址。
        periods: 时间段可迭代对象，每个元素可解包为 (start, end)。
            会在启动 Excel 前物化并校验；同批次的开始年月必须唯一。
        output_dir: 输出目录，调用方需确保已创建且专用于本次调用；装饰器会处理
            该目录直接包含的所有非 .zip 文件

    Returns:
        输出目录路径（经 @auto_zip 装饰器后处理——压缩 .xls 为 .zip 并删除原文件）

    Raises:
        ValueError: 同批次周期生成重复输出文件名(ERR_OUTPUT_PATH)。
        FileNotFoundError: 模板文件不存在(ERR_TEMPLATE_NOT_FOUND)。
        RuntimeError: 批量压缩失败(ERR_ARCHIVE_PARTIAL)，或 Excel 处理失败。

    Excel 处理失败时会尽力删除本次调用新建的部分 .xls 文件；清理失败不会替换原始
    异常。调用前已存在的同名文件不会在回滚时删除，但成功保存同名输出时仍会覆盖它。
    后续压缩阶段不是事务；若部分压缩失败，已成功的压缩包和未处理的 .xls 可能并存。

    模板选择规则：
        - year < 2019  → 2018.xls(旧个税)
        - year >= 2019 → 2019.xls(新个税)
    """
    contexts = [{**data, 'start': s, 'end': e} for s, e in periods]
    output_names: set[str] = set()
    for context in contexts:
        name = f"{context['start'].year}_{context['start'].month:02}.xls"
        if name in output_names:
            raise ValueError(
                format_user_error(
                    ERR_OUTPUT_PATH,
                    f"Duplicate period output filename: {name}",
                    f"多个周期对应同一个输出文件：{name}",
                )
            )
        output_names.add(name)
    generated_paths: list[Path] = []

    try:
        with xw.App(visible=False, add_book=False) as app:
            app.display_alerts = False
            app.screen_updating = False

            for context in contexts:
                year, month = context['start'].year, context['start'].month
                template_year = 2018 if year < 2019 else 2019

                template_path = search_template_file_cached(
                    TEMPLATE_NAMES[template_year], suffix='xls'
                )
                if template_path is None:
                    pattern = f"{TEMPLATE_NAMES[template_year]}.xls"
                    raise FileNotFoundError(
                        format_user_error(
                            ERR_TEMPLATE_NOT_FOUND,
                            f"Template not found: {ascii_safe(pattern)}",
                            f"找不到模板文件: {pattern}",
                        )
                    )
                wb = app.books.open(template_path)
                output_path = output_dir / f'{year}_{month:02}.xls'
                # 失败时只回滚本次调用新建的文件，不删除调用方已有文件或符号链接。
                if not output_path.exists() and not output_path.is_symlink():
                    generated_paths.append(output_path)
                try:
                    sht = wb.sheets[0]

                    for key, value in context.items():
                        sht.range(key).value = value

                    wb.save(output_path)
                except BaseException:
                    # 清理失败只能记日志，不能替换正在传播的写入/保存异常。
                    try:
                        wb.close()
                    except Exception:
                        logger.warning("Failed to close workbook after generation error", exc_info=True)
                    raise
                else:
                    # 没有先前异常时，关闭失败仍应作为本次失败向上传播。
                    wb.close()
    except Exception:
        for path in generated_paths:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass
        raise

    return output_dir
