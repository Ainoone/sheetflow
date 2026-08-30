"""列式到行式数据转置模块。

将列式映射转置为行式记录的可迭代容器，配合 docxtpl 批量渲染 Word 报告。
构造时会浅复制键和列，后续迭代不受调用方修改映射或列表的影响。

典型用例：
    >>> data = {'company': ['Apple', 'Microsoft'], 'price': [999.0, 899.0]}
    >>> rs = RecordSet(data)
    >>> for record in rs:
    ...     print(record)
    {'company': 'Apple', 'price': 999.0}
    {'company': 'Microsoft', 'price': 899.0}

    >>> # 随机访问
    >>> rs.records[0]
    {'company': 'Apple', 'price': 999.0}
"""

from collections.abc import Iterator, Mapping
from typing import Any

from .user_errors import (
    ERR_RECORD_COLUMN_TYPE,
    ERR_RECORD_DATA_TYPE,
    ERR_RECORD_KEY_TYPE,
    ERR_RECORD_LENGTH_MISMATCH,
    ascii_safe,
    format_user_error,
)


class RecordSet:
    """列式到行式转置容器，支持迭代和随机访问。

    将列式映射（每个字符串键对应一个列表）转换为行式记录的迭代器。
    所有列表必须长度一致。

    Args:
        data: 列式映射，格式为 {列名: [值列表]}
        dropna: 若为 True,迭代时自动跳过含 None 值的行

    Raises:
        TypeError: data 不是映射、列名不是字符串或列值不是列表时抛出
        ValueError: 当列表长度不一致时抛出

    Examples:
        >>> # 基本用法
        >>> data = {'name': ['Alice', 'Bob'], 'age': [25, 30]}
        >>> rs = RecordSet(data)
        >>> list(rs)
        [{'name': 'Alice', 'age': 25}, {'name': 'Bob', 'age': 30}]

        >>> # 随机访问
        >>> rs.records[0]
        {'name': 'Alice', 'age': 25}

        >>> # 多次迭代
        >>> for record in rs:
        ...     print(record['name'])
        Alice
        Bob

        >>> # dropna 跳过含 None 的行
        >>> rs = RecordSet({'name': ['Alice', None], 'age': [25, 30]}, dropna=True)
        >>> list(rs)
        [{'name': 'Alice', 'age': 25}]

    Note:
        - 迭代器可多次调用，每次都会重新生成
        - 构造时会浅复制键和列；不会深复制列表中的单元格对象
        - records 属性每次返回新的行式列表，修改结果不会影响记录集
        - 空字典输入会产生长度为 0 的容器（迭代不产生任何记录）
        - 构造时会验证所有列表长度一致，不一致会抛出 ValueError
    """

    def __init__(
        self,
        data: Mapping[str, list[Any]],
        dropna: bool = False,
    ) -> None:
        """初始化记录集。

        Args:
            data: 列式映射，每个字符串键对应一个列表
            dropna: 若为 True,迭代时自动跳过含 None 值的行

        Raises:
            TypeError: data 不是映射、列名不是字符串或列值不是列表时抛出
            ValueError: 当列表长度不一致时抛出

        Examples:
            >>> rs = RecordSet({'name': ['Alice'], 'age': [25]})
            >>> list(rs)
            [{'name': 'Alice', 'age': 25}]

            >>> # 长度不一致会抛出异常
            >>> RecordSet({'name': ['Alice', 'Bob'], 'age': [25]})  # doctest: +ELLIPSIS
            Traceback (most recent call last):
                ...
            ValueError: ERR_RECORD_LENGTH_MISMATCH: All column lists must have ...

            >>> # dropna 跳过含 None 的行
            >>> rs = RecordSet({'name': ['Alice', None], 'age': [25, 30]}, dropna=True)
            >>> list(rs)
            [{'name': 'Alice', 'age': 25}]
        """
        if not isinstance(data, Mapping):
            raise TypeError(
                format_user_error(
                    ERR_RECORD_DATA_TYPE,
                    f"RecordSet data must be a mapping; type={type(data).__name__}",
                    f"RecordSet 数据必须是映射。当前类型: {type(data).__name__}",
                )
            )

        items = tuple(data.items())
        for key, column in items:
            if not isinstance(key, str):
                raise TypeError(
                    format_user_error(
                        ERR_RECORD_KEY_TYPE,
                        f"RecordSet column names must be strings; key={ascii_safe(key)}; "
                        f"type={type(key).__name__}",
                        f"RecordSet 列名必须是字符串。当前列名: {key}",
                    )
                )
            if not isinstance(column, list):
                raise TypeError(
                    format_user_error(
                        ERR_RECORD_COLUMN_TYPE,
                        f"RecordSet columns must be lists; column={ascii_safe(key)}; "
                        f"type={type(column).__name__}",
                        f"RecordSet 列必须是列表。当前列: {key}",
                    )
                )

        lengths = {key: len(column) for key, column in items}
        unique_lengths = set(lengths.values())
        if len(unique_lengths) > 1:
            raise ValueError(
                format_user_error(
                    ERR_RECORD_LENGTH_MISMATCH,
                    f"All column lists must have the same length; lengths={ascii_safe(lengths)}",
                    f"所有列表长度必须一致。当前长度: {lengths}",
                )
            )

        self._keys = tuple(key for key, _ in items)
        self._columns = tuple(tuple(column) for _, column in items)
        self._dropna = dropna

    def __iter__(self) -> Iterator[dict]:
        """迭代产出行式记录，若 dropna=True 则跳过含 None 值的行。

        Yields:
            dict: 每行数据的字典表示，键为列名，值为对应单元格的值

        Examples:
            >>> data = {'name': ['Alice', 'Bob'], 'age': [25, 30]}
            >>> rs = RecordSet(data)
            >>> for record in rs:
            ...     print(record)
            {'name': 'Alice', 'age': 25}
            {'name': 'Bob', 'age': 30}

            >>> # dropna 跳过含 None 的行
            >>> rs = RecordSet({'name': ['Alice', None], 'age': [25, 30]}, dropna=True)
            >>> list(rs)
            [{'name': 'Alice', 'age': 25}]

        Note:
            可多次调用，每次都会重新生成迭代器
        """
        for row in zip(*self._columns):
            if self._dropna and any(value is None for value in row):
                continue
            yield dict(zip(self._keys, row))

    def __repr__(self) -> str:
        """返回对象的字符串表示。

        Returns:
            包含键名的字符串，便于调试

        Examples:
            >>> rs = RecordSet({'name': ['Alice'], 'age': [25]})
            >>> repr(rs)
            "RecordSet(keys=['name', 'age'])"
        """
        return f"RecordSet(keys={list(self._keys)})"

    @property
    def records(self) -> list[dict]:
        """完整行式列表，需随机访问时使用。

        Returns:
            所有记录的列表，每个元素是一个字典

        Examples:
            >>> data = {'name': ['Alice', 'Bob'], 'age': [25, 30]}
            >>> rs = RecordSet(data)
            >>> rs.records
            [{'name': 'Alice', 'age': 25}, {'name': 'Bob', 'age': 30}]

            >>> # 支持索引访问
            >>> rs.records[0]
            {'name': 'Alice', 'age': 25}

            >>> # 支持切片
            >>> rs.records[0:1]
            [{'name': 'Alice', 'age': 25}]

        Note:
            - 每次访问都返回新的列表和记录字典
            - 修改返回结果不会影响记录集或后续访问
            - 如果只需要遍历一次，直接迭代对象更节省内存
        """
        return list(self)
