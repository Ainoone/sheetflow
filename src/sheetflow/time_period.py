"""时间段切分工具模块。

将时间段按指定频率（月/季/年）切分为自然周期区间，输出迭代器便于写入 Excel/Word。
输入日期可为 datetime.date、datetime.datetime 或 (year, month, day) 元组；
输出统一为 datetime.date。公共入口会在调用时立即转换并校验输入，周期本身按需生成。

典型用例：
    >>> from datetime import date
    >>> ts = {"start": date(2024, 1, 15), "end": date(2024, 3, 20), "freq": "M"}
    >>> periods = list(generate_period_range(ts))
    >>> periods[0]
    Period(start=datetime.date(2024, 1, 1), end=datetime.date(2024, 1, 31))

支持的频率：
    - "M": 月度切分
    - "Q": 季度切分
    - "Y": 年度切分
    - "N": 无频率（整段作为单个区间）

依赖：Python 标准库 datetime/calendar
"""

import calendar
import datetime
from collections.abc import Iterator, Mapping
from typing import NamedTuple

NO_FREQ = "N"

FREQS = {"M", "Q", "Y"}
_SUPPORTED_FREQS = FREQS | {NO_FREQ}


class Period(NamedTuple):
    """起止日期均包含在内的时间区间。"""

    start: datetime.date
    end: datetime.date


def _as_date(value: object) -> datetime.date:
    """将支持的日期输入统一转换为 datetime.date。

    Raises:
        TypeError: value 不是 date、datetime 或三元素日期元组。
        ValueError: 日期元组不能构造有效日期。
    """
    # datetime 是 date 的子类，必须在 date 之前判断以去除时间部分。
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, tuple) and len(value) == 3:
        return datetime.date(*value)

    raise TypeError(
        "Date value must be date, datetime, or a (year, month, day) tuple"
    )


def _natural_period(day: datetime.date, freq: str) -> Period:
    """返回 day 所在的完整自然月、自然季度或自然年。

    freq 由公共入口校验，此处只接收 "M"、"Q" 或 "Y"。
    """
    if freq == "M":
        first_month = last_month = day.month
    elif freq == "Q":
        first_month = ((day.month - 1) // 3) * 3 + 1
        last_month = first_month + 2
    else:  # freq == "Y"，公共入口已校验频率。
        first_month = 1
        last_month = 12

    start = datetime.date(day.year, first_month, 1)
    end = datetime.date(
        day.year,
        last_month,
        calendar.monthrange(day.year, last_month)[1],
    )
    return Period(start, end)


def _iter_natural_periods(
    start: datetime.date,
    end: datetime.date,
    freq: str,
) -> Iterator[Period]:
    """按已校验频率生成覆盖有序闭区间 [start, end] 的连续自然周期。"""
    current = start

    while True:
        period = _natural_period(current, freq)
        yield period

        # 先判断是否已覆盖终点，避免在 date.max 后推进到不存在的年份 10000。
        if period.end >= end:
            return

        current = period.end + datetime.timedelta(days=1)


def generate_period_range(timeseries: Mapping[str, object]) -> Iterator[Period]:
    """根据时间序列映射生成周期区间。

    调用时立即转换日期、校验频率并规范化日期顺序；返回的迭代器按需生成周期，
    因而输入错误会在调用方创建输出目录或启动 Excel 之前暴露。

    Args:
        timeseries: 包含 "start" / "end" / "freq" 键的映射
            - start: datetime.date、datetime.datetime 或 tuple(year, month, day)
            - end: datetime.date、datetime.datetime 或 tuple(year, month, day)
            - freq: "M"（月）/ "Q"（季）/ "Y"（年）/ "N"（不分段）

    Returns:
        Period 迭代器；每个区间的起止值均为 datetime.date。start != end 时，
        M/Q/Y 返回覆盖输入范围的完整自然周期；N 返回规范化为正序但不切分的
        原始区间。start == end 时所有有效频率都返回该单日区间。

    Examples:
        >>> from datetime import date
        >>> ts = {"start": date(2024, 1, 15), "end": date(2024, 3, 20), "freq": "M"}
        >>> list(generate_period_range(ts)) == [
        ...     Period(date(2024, 1, 1), date(2024, 1, 31)),
        ...     Period(date(2024, 2, 1), date(2024, 2, 29)),
        ...     Period(date(2024, 3, 1), date(2024, 3, 31)),
        ... ]
        True

        >>> ts = {"start": date(2024, 1, 1), "end": date(2024, 12, 31), "freq": "N"}
        >>> list(generate_period_range(ts))
        [Period(start=datetime.date(2024, 1, 1), end=datetime.date(2024, 12, 31))]

    Raises:
        TypeError: start 或 end 不是支持的日期类型
        ValueError: freq 无效，或日期元组不能构造有效日期
        KeyError: 字典缺少必需键时抛出

    Note:
        日期转换后始终先校验 freq；有效频率下，start == end 时返回该单日区间，
        不对齐到周期边界。start > end 时统一交换为正序。
    """
    # 公共入口不使用 yield：立即校验，避免下游副作用发生后才暴露输入错误。
    start = _as_date(timeseries["start"])
    end = _as_date(timeseries["end"])
    freq = timeseries["freq"]

    if not isinstance(freq, str) or freq not in _SUPPORTED_FREQS:
        raise ValueError(
            f"Invalid freq {freq!r}. Must be one of: {sorted(_SUPPORTED_FREQS)}"
        )

    if start > end:
        start, end = end, start

    if freq == NO_FREQ or start == end:
        return iter((Period(start, end),))

    return _iter_natural_periods(start, end, freq)
