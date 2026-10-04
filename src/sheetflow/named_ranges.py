"""Excel 命名区域(Named Ranges)处理工具模块。

本模块提供读取和批量创建 Excel 工作表级命名区域的功能，基于 xlwings 库实现。
主要用于自动化场景中的命名区域管理，支持从工作表读取命名区域数据，
以及根据映射字典批量创建命名区域。

读取侧刻意使用 sheet.names,而不是 book.names,以确保只处理当前工作表的
工作表级命名区域，不混入工作簿全局名称。

Public API
----------
NamedRangeDict                  — 读取工作表命名区域为字典
build_named_range_map           — 从双列表构建 {名称: 地址} 映射
create_sheet_named_ranges       — 批量创建工作表级命名区域(O(n) COM 调用）
create_named_ranges_from_dict   — 便捷封装(build + create)

典型用法
--------
读取命名区域：
    >>> sheet = xw.sheets.active
    >>> nrd = NamedRangeDict(sheet)
    >>> data = nrd.data  # {'company': '某公司', 'amount': 10000, ...}
    >>> output_dir = str(Path(nrd.sheet.book.fullname).parent)

批量创建命名区域：
    >>> mapping = {'company': 'A1', 'amount': 'B1'}
    >>> result = create_sheet_named_ranges(sheet, mapping)
    >>> print(result)  # {'ok': ['company', 'amount'], 'failed': []}

从双列表创建：
    >>> raw = {'Named': ['company', 'amount'], 'Range': ['A1', 'B1']}
    >>> result = create_named_ranges_from_dict(sheet, raw)
"""

import datetime
import logging
import xlwings as xw
from more_itertools import always_iterable

from .user_errors import (
    ERR_NAMED_RANGE_MAP,
    ERR_NAMED_RANGE_RESTORE_FAILED,
    ascii_safe,
    format_user_error,
)

logger = logging.getLogger(__name__)


def _extract_local_name(raw: str) -> str:
    """从 xlwings 工作表级完整名称中提取不含工作表前缀的名称本体。

    ``sheet.names`` 返回的名称带有 ``SheetName!`` 前缀。命名区域本体不含
    ASCII ``!``，因此从右侧分隔可以兼容工作表名自身包含 ``!`` 的情况。
    """
    _, _, local_name = raw.rpartition("!")
    return local_name


class NamedRangeDict:
    """将工作表命名区域构建为 {名称: 值} 字典。

    采用懒加载策略，首次访问 data 属性时才执行 COM 调用读取命名区域，
    避免不必要的性能开销。自动提取工作表级名称本体（如
    ``'Sheet1'!name`` → ``name``）。读取范围限定为传入 sheet 的
    sheet.names,不读取工作簿级全局名称。

    Attributes:
        sheet (xw.Sheet): 关联的 xlwings 工作表对象（只读）
        data (dict): {不含工作表前缀的名称: 值} 字典，懒加载

    Examples:
        直接实例化：
            >>> sheet = xw.sheets['Sheet1']
            >>> nrd = NamedRangeDict(sheet)
            >>> print(nrd.data)
            {'company': '某公司', 'date': datetime.date(2026, 5, 2)}

        从 Excel 按钮调用，获取工作簿路径：
            >>> nrd = NamedRangeDict.from_caller()
            >>> output_dir = str(Path(nrd.sheet.book.fullname).parent)
            >>> data = nrd.data
    """

    def __init__(self, sheet: xw.Sheet):
        """初始化命名区域字典对象。

        Args:
            sheet: xlwings 工作表对象
        """
        self._sheet = sheet
        self._data: dict | None = None

    @property
    def sheet(self) -> xw.Sheet:
        """关联的 xlwings 工作表对象（只读）。"""
        return self._sheet

    @property
    def data(self) -> dict:
        """返回 {不含工作表前缀的名称: 值} 字典（懒加载）。"""
        if self._data is None:
            self._data = self._build()
        return self._data

    @classmethod
    def from_caller(cls) -> "NamedRangeDict":
        """通过 Excel 按钮调用时使用的便捷构造方法。

        从 xlwings 的 caller 上下文获取当前活动工作表并实例化。
        适用于在 Excel 中通过 "Run main" 按钮或宏触发的场景。

        Returns:
            NamedRangeDict: 基于调用者活动工作表的实例

        Examples:
            在 Excel VBA 宏中调用的 Python 函数：
                >>> def main():
                ...     nrd = NamedRangeDict.from_caller()
                ...     output_dir = str(Path(nrd.sheet.book.fullname).parent)
                ...     process_data(nrd.data, output_dir)
        """
        sheet = xw.Book.caller().sheets.active
        return cls(sheet)

    def _build(self) -> dict:
        """构建命名区域字典（内部方法）。

        遍历当前工作表 sheet.names 中的所有工作表级命名区域，执行 COM 调用获取值。
        每个 _get_value() 调用对应一次 COM 往返，懒加载确保仅执行一次。

        Returns:
            dict: {不含工作表前缀的名称: 值} 字典
        """
        # 每次取值都会触发一次 COM 往返；延迟构建确保同一实例只读取一次。
        return {
            _extract_local_name(name_obj.name): self._get_value(name_obj)
            for name_obj in self._sheet.names
        }

    @staticmethod
    def _get_value(name_obj):
        """获取命名区域的值，应用类型转换选项。

        Args:
            name_obj: xlwings Name 对象

        Returns:
            命名区域的值，日期转为 datetime.date,空单元格转为 None
        """
        return name_obj.refers_to_range.options(
            dates=datetime.date,
            empty=None,
        ).value


def _as_config_list(value) -> list:
    """将 xlwings 的单单元格标量和多单元格列表统一为列表。"""
    if value is None:
        return []
    # xlwings 对单单元格返回标量；统一包装避免把 "CN" 按字符拆分。
    return list(always_iterable(value))


def build_named_range_map(namedrange_dict: dict) -> dict:
    """从原始双键字典构建 {名称: 地址} 映射，过滤任意一方为 None 的条目。

    将 {'Named': [...], 'Range': [...]} 格式的原始数据转换为
    {名称: 地址} 映射字典，自动过滤掉名称或地址为 None 的无效条目。

    Args:
        namedrange_dict: 包含 'Named' 和 'Range' 键的字典
            - 'Named': 单个名称或名称列表
            - 'Range': 单个地址或对应的单元格地址列表

    Returns:
        dict: {名称: 地址} 映射，已过滤 None 值

    Raises:
        ValueError: Named 和 Range 列表长度不匹配、名称不是字符串，或存在不区分大小写的重复名称
            (ERR_NAMED_RANGE_MAP)。

    Examples:
        >>> raw = {'Named': ['company', 'amount', None], 'Range': ['A1', 'B1', 'C1']}
        >>> build_named_range_map(raw)
        {'company': 'A1', 'amount': 'B1'}
    """
    names = _as_config_list(namedrange_dict.get("Named"))
    ranges = _as_config_list(namedrange_dict.get("Range"))

    if len(names) != len(ranges):
        raise ValueError(
            format_user_error(
                ERR_NAMED_RANGE_MAP,
                f"Named/Range length mismatch: {len(names)} names vs {len(ranges)} ranges. "
                "Check the source sheet for missing or extra entries.",
                f"Named/Range 长度不一致：{len(names)} 个名称 vs {len(ranges)} 个地址。"
                "请检查源工作表是否存在缺失或多余条目。",
            )
        )

    result = {}
    seen_names: dict[str, str] = {}
    for name, addr in zip(names, ranges):
        if name is None or addr is None:
            continue

        if not isinstance(name, str):
            raise ValueError(
                format_user_error(
                    ERR_NAMED_RANGE_MAP,
                    f"Named range names must be strings; name={ascii_safe(name)}",
                    f"命名区域名称必须为文本：{name}",
                )
            )
        normalized_name = name.casefold()
        previous = seen_names.get(normalized_name)
        if previous is not None:
            raise ValueError(
                format_user_error(
                    ERR_NAMED_RANGE_MAP,
                    "Duplicate named ranges are not allowed; "
                    f"names={ascii_safe(previous)}, {ascii_safe(name)}",
                    f"命名区域名称重复：{previous}、{name}",
                )
            )
        seen_names[normalized_name] = name

        result[name] = addr

    return result


def create_sheet_named_ranges(
    sheet: xw.Sheet,
    named_range_map: dict,
    overwrite: bool = True,
) -> dict:
    """批量创建工作表级命名区域，单个失败不中断整体流程；覆盖已有定义时，
    新定义创建失败会尝试恢复旧定义。

    采用 O(n) 算法一次性读取所有现有命名区域，避免逐个查询的 O(n²) 开销。
    覆盖模式按不含工作表前缀的名称本体建立索引，因此 Excel/xlwings 读回的
    工作表前缀是否带引号不影响匹配；索引使用 ``casefold``，以符合 Excel
    定义名称不区分大小写的行为。创建新名称时统一按 Excel 引用语法为工作表名
    加单引号，并转义工作表名内部的英文单引号；这种写法对普通工作表名和包含
    空格、括号或其他标点的工作表名都有效。单个命名区域创建失败时，记录日志并
    继续处理剩余项，最终返回成功和失败的名称列表。

    Args:
        sheet: 目标 xlwings 工作表对象
        named_range_map: {名称: 单元格地址} 映射字典
            - 名称: 命名区域的名称（不含工作表前缀）
            - 地址: 单元格地址，如 'A1', 'B2:C5'
        overwrite: 是否覆盖同名定义，默认 True
            - True: 先删除同名旧定义再创建；前缀引号和名称大小写不影响匹配
            - False: 保留现有定义，可能导致创建失败

    Returns:
        dict: 包含两个键的结果字典
            - 'ok': 成功创建的名称列表
            - 'failed': 创建失败的名称列表

    Raises:
        RuntimeError: 覆盖已有命名区域时，新定义创建失败且旧定义也无法恢复
            (ERR_NAMED_RANGE_RESTORE_FAILED)。

    Note:
        此函数不是事务。普通创建失败会继续处理剩余名称，已成功的更改保留在
        工作簿内存中；调用方根据返回结果决定是否保存工作簿。

    Examples:
        >>> sheet = xw.sheets['Sheet1']
        >>> mapping = {'company': 'A1', 'amount': 'B1', 'invalid': 'Z999'}
        >>> result = create_sheet_named_ranges(sheet, mapping)
        >>> print(result)
        {'ok': ['company', 'amount'], 'failed': ['invalid']}

        工作表名包含括号或空格时也可创建：
        >>> sheet = xw.sheets['扣缴个人所得税申报表（适用于综合所得预扣预缴）']
        >>> create_sheet_named_ranges(sheet, {'CN': 'A1'})
        {'ok': ['CN'], 'failed': []}

        不覆盖现有定义：
        >>> result = create_sheet_named_ranges(sheet, mapping, overwrite=False)
    """
    result: dict[str, list] = {"ok": [], "failed": []}

    # sheet.names 已限定工作表范围；casefold 匹配 Excel 名称不区分大小写的规则。
    existing: dict[str, xw.Name] = (
        {
            _extract_local_name(name_obj.name).casefold(): name_obj
            for name_obj in sheet.names
        }
        if overwrite
        else {}
    )

    # 工作表名统一加引号并转义单引号，避免逐项猜测 Excel 引用规则。
    safe_name = sheet.name.replace("'", "''")
    quoted_sheet = f"'{safe_name}'"

    for name, addr in named_range_map.items():
        full_name = f"{quoted_sheet}!{name}"
        old_name = None
        old_name_text = None
        old_refers_to = None
        old_deleted = False
        try:
            old_name = existing.get(name.casefold()) if overwrite else None
            if old_name:
                old_name_text = old_name.name
                old_refers_to = old_name.refers_to
                old_name.delete()
                old_deleted = True

            sheet.range(addr).name = full_name
            result["ok"].append(name)

        except Exception as e:
            if old_deleted and old_name_text and old_refers_to:
                try:
                    # Worksheet.Names.Add 接受裸名；Name.name 返回的表级名称含前缀。
                    sheet.names.add(
                        _extract_local_name(old_name_text), old_refers_to
                    )
                except Exception as restore_error:
                    logger.error(
                        f"[恢复失败] {name} ({addr}): {restore_error}"
                    )
                    raise RuntimeError(
                        format_user_error(
                            ERR_NAMED_RANGE_RESTORE_FAILED,
                            f"Unable to restore named range: {old_name_text}; "
                            f"error={ascii_safe(restore_error)}",
                            f"无法恢复命名区域：{old_name_text}；错误：{restore_error}",
                        )
                    ) from restore_error
            logger.warning(f"[跳过] {name} ({addr}): {e}")
            result["failed"].append(name)

    return result


def create_named_ranges_from_dict(
    sheet: xw.Sheet,
    namedrange_dict: dict,
    overwrite: bool = True,
) -> dict:
    """从双列表字典创建命名区域的便捷封装函数。

    串联 build_named_range_map 和 create_sheet_named_ranges 两个步骤，
    提供一站式接口从原始 {'Named': [...], 'Range': [...]} 格式直接创建命名区域。

    Args:
        sheet: 目标 xlwings 工作表对象
        namedrange_dict: 包含 'Named' 和 'Range' 键的字典
            - 'Named': 名称列表
            - 'Range': 对应的单元格地址列表
        overwrite: 是否覆盖同名定义，默认 True

    Returns:
        dict: 包含两个键的结果字典
            - 'ok': 成功创建的名称列表
            - 'failed': 创建失败的名称列表

    Raises:
        ValueError: Named 和 Range 列表长度不匹配、名称不是字符串，或存在重复名称
            (ERR_NAMED_RANGE_MAP，由 build_named_range_map 抛出)。
        RuntimeError: 覆盖已有命名区域时，新定义创建失败且旧定义也无法恢复
            (ERR_NAMED_RANGE_RESTORE_FAILED，由 create_sheet_named_ranges 抛出)。

    Examples:
        >>> sheet = xw.sheets['Sheet1']
        >>> raw = {
        ...     'Named': ['company', 'amount', 'date'],
        ...     'Range': ['A1', 'B1', 'C1']
        ... }
        >>> result = create_named_ranges_from_dict(sheet, raw)
        >>> print(f"成功: {len(result['ok'])}, 失败: {len(result['failed'])}")
        成功: 3, 失败: 0
    """
    named_range_map = build_named_range_map(namedrange_dict)
    return create_sheet_named_ranges(sheet, named_range_map, overwrite)
