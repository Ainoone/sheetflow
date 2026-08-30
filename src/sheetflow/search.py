"""模板文件搜索模块。

提供模板文件的递归搜索功能，支持自定义搜索路径和只保存有效成功结果的有界缓存。

典型用例：
    >>> from pathlib import Path
    >>> # 在默认模板目录搜索
    >>> template = search_template_file('个税申报表')
    >>> print(template)
    PosixPath('/path/to/Template/Word/个税申报表.docx')

    >>> # 在自定义路径搜索
    >>> template = search_template_file('合同模板', suffix='docx', search_path='/custom/path')

    >>> # 使用缓存版本
    >>> template = search_template_file_cached('个税申报表')

常量：
    BASE_DIR: 模块所在目录的绝对路径
    TEMPLATE_DIR: 默认模板目录路径 (BASE_DIR / 'Template')
"""

from functools import lru_cache
from pathlib import Path
import logging

from .user_errors import (
    ERR_SEARCH_PATH,
    ERR_TEMPLATE_NAME,
    ERR_TEMPLATE_NOT_FOUND,
    ascii_safe,
    format_user_error,
)

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent.resolve()
TEMPLATE_DIR = BASE_DIR / "Template"


def _validate_template_name(name: str) -> None:
    """确保模板名是纯文件名，而不是可穿越目录的路径。"""
    # 同时拒绝 POSIX 和 Windows 分隔符，避免校验结果随运行平台变化。
    if (
        not isinstance(name, str)
        or not name
        or name in {".", ".."}
        or "/" in name
        or "\\" in name
    ):
        raise ValueError(
            format_user_error(
                ERR_TEMPLATE_NAME,
                f"Template name must not contain a path; name={ascii_safe(name)}",
                f"模板名称不能包含路径：{name}",
            )
        )


def search_template_file(
    name: str,
    suffix: str = "docx",
    search_path: str | Path | None = None
) -> Path | None:
    """在指定路径中递归搜索模板文件。

    Args:
        name: 文件名（不含扩展名），例如 '个税申报表'。传入含扩展名的值会
            组成重复扩展名的搜索模式。
        suffix: 文件扩展名，默认 'docx'
        search_path: 搜索路径，默认使用 TEMPLATE_DIR。支持 str 或 Path 对象

    Returns:
        找到的文件路径(Path 对象），未找到返回 None

    Raises:
        ValueError: 模板名包含路径(ERR_TEMPLATE_NAME)，或搜索路径不存在或不是目录
            (ERR_SEARCH_PATH)。

    Examples:
        >>> # 在默认模板目录搜索
        >>> result = search_template_file('个税申报表')
        >>> print(result)
        PosixPath('/path/to/Template/Word/个税申报表.docx')

        >>> # 搜索 Excel 模板
        >>> result = search_template_file('工资表', suffix='xlsx')

        >>> # 在自定义路径搜索
        >>> result = search_template_file('合同', search_path='/custom/templates')

        >>> # 文件不存在时返回 None
        >>> result = search_template_file('不存在的文件')
        >>> print(result)
        None

    Note:
        - 递归搜索所有子目录，并按完整文件名精确匹配
        - 未找到文件时会记录警告日志
        - 此函数不使用缓存，每次调用都会重新搜索
        - 如需缓存，使用 search_template_file_cached()
    """
    path = Path(search_path) if search_path else TEMPLATE_DIR
    _validate_template_name(name)

    if not path.exists():
        raise ValueError(
            format_user_error(
                ERR_SEARCH_PATH,
                f"Search path does not exist: {ascii_safe(path)}",
                f"搜索路径不存在: {path}",
            )
        )
    if not path.is_dir():
        raise ValueError(
            format_user_error(
                ERR_SEARCH_PATH,
                f"Search path is not a directory: {ascii_safe(path)}",
                f"搜索路径不是目录: {path}",
            )
        )

    pattern = f"{name}.{suffix}"

    # 按完整文件名比较，避免把模板名或后缀中的 glob 元字符解释为搜索语法。
    result = next(
        (
            candidate
            for candidate in path.rglob("*")
            if candidate.is_file() and candidate.name == pattern
        ),
        None,
    )

    if result is None:
        logger.warning(f"未找到模板文件: {pattern} (搜索路径: {path})")
    else:
        logger.debug(f"找到模板文件: {result}")

    return result


def require_template(
    name: str,
    suffix: str = "docx",
    search_path: str | Path | None = None,
) -> Path:
    """搜索模板文件，未找到时抛出 FileNotFoundError。

    Args:
        name: 文件名（不含扩展名），例如 '个税申报表'。
        suffix: 文件扩展名，默认 'docx'。
        search_path: 搜索路径，默认使用 TEMPLATE_DIR。

    Returns:
        找到的模板文件路径。

    Raises:
        FileNotFoundError: 未找到指定模板文件(ERR_TEMPLATE_NOT_FOUND)。
        ValueError: 模板名包含路径(ERR_TEMPLATE_NAME)，或搜索路径不存在或不是目录
            (ERR_SEARCH_PATH)。
    """
    template_path = search_template_file(name, suffix, search_path)
    if template_path is None:
        pattern = f"{name}.{suffix}"
        raise FileNotFoundError(
            format_user_error(
                ERR_TEMPLATE_NOT_FOUND,
                f"Template not found: {ascii_safe(pattern)}",
                f"找不到模板文件: {pattern}",
            )
        )
    return template_path


class _TemplateCacheMiss(LookupError):
    """内部信号：未命中不应进入 lru_cache。"""


@lru_cache(maxsize=128)
def _search_template_file_positive_cached(
    name: str,
    suffix: str = "docx",
    search_path: str | Path | None = None
) -> Path:
    """只把成功搜索到的模板加入有界缓存。"""
    result = search_template_file(name, suffix, search_path)
    if result is None:
        # functools.lru_cache 不缓存异常，因此运行时新增模板可在下次调用被发现。
        raise _TemplateCacheMiss
    return result


def search_template_file_cached(
    name: str,
    suffix: str = "docx",
    search_path: str | Path | None = None
) -> Path | None:
    """缓存版本的模板文件搜索。

    使用 functools.lru_cache 缓存最近 128 组成功搜索，提高重复查询的性能。
    未找到结果不会缓存；已缓存路径失效时会清除缓存并立即重新搜索。

    Args:
        name: 文件名（不含扩展名）
        suffix: 文件扩展名，默认 'docx'
        search_path: 搜索路径，默认使用 TEMPLATE_DIR

    Returns:
        找到的文件路径，未找到返回 None

    Raises:
        ValueError: 模板名包含路径(ERR_TEMPLATE_NAME)，或搜索路径不存在或不是目录
            (ERR_SEARCH_PATH)。

    Examples:
        >>> # 第一次调用会搜索文件系统
        >>> result = search_template_file_cached('个税申报表')
        >>> # 第二次调用直接返回缓存结果
        >>> result = search_template_file_cached('个税申报表')

        >>> # 清除缓存
        >>> search_template_file_cached.cache_clear()

    Note:
        - 缓存基于参数值，只保存仍存在的成功结果
        - 未找到结果不进入缓存，运行时添加模板后下一次调用即可发现
        - 已缓存文件被移动或删除时会自动重新搜索
        - 最多缓存 128 组参数，避免 Excel 长进程中的无界增长
        - 使用 search_template_file_cached.cache_clear() 清除所有缓存
        - 使用 search_template_file_cached.cache_info() 查看缓存统计
    """
    try:
        result = _search_template_file_positive_cached(name, suffix, search_path)
    except _TemplateCacheMiss:
        return None

    if result.is_file():
        return result

    # lru_cache 不支持按键删除；失效路径很少见，清空有界缓存后重搜最可靠。
    _search_template_file_positive_cached.cache_clear()
    try:
        return _search_template_file_positive_cached(name, suffix, search_path)
    except _TemplateCacheMiss:
        return None


# cache_clear/cache_info 不在包装器的 __dict__ 中，functools.wraps 无法转发；
# 显式保留原公开函数已文档化的缓存管理接口。
search_template_file_cached.cache_clear = (  # type: ignore[attr-defined]
    _search_template_file_positive_cached.cache_clear
)
search_template_file_cached.cache_info = (  # type: ignore[attr-defined]
    _search_template_file_positive_cached.cache_info
)
search_template_file_cached.cache_parameters = (  # type: ignore[attr-defined]
    _search_template_file_positive_cached.cache_parameters
)
