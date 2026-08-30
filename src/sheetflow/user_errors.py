"""用户可见错误码与错误文本的唯一权威入口。

本模块只依赖 Python 标准库，供上层流程和底层路径工具共同使用，避免通过
``common`` 形成反向依赖或复制 ASCII-safe 实现。

错误码是 Excel/VBA 链路依赖的用户可见契约。新增或删除生产错误码时，必须
同步更新 README 错误码表；测试会校验注册表与生产代码使用集合完全一致。
"""

from typing import Any

ERROR_PREFIX = "ERR_"

ERR_ARCHIVE_PARTIAL = "ERR_ARCHIVE_PARTIAL"
ERR_FILE_NOT_FOUND = "ERR_FILE_NOT_FOUND"
ERR_INVALID_PATH = "ERR_INVALID_PATH"
ERR_MISSING_NAMED_RANGE = "ERR_MISSING_NAMED_RANGE"
ERR_MIXED_RECORD_SHAPE = "ERR_MIXED_RECORD_SHAPE"
ERR_NAMED_RANGE_MAP = "ERR_NAMED_RANGE_MAP"
ERR_NAMED_RANGE_PARTIAL = "ERR_NAMED_RANGE_PARTIAL"
ERR_NAMED_RANGE_RESTORE_FAILED = "ERR_NAMED_RANGE_RESTORE_FAILED"
ERR_NOT_A_FILE = "ERR_NOT_A_FILE"
ERR_NO_NAMED_RANGES = "ERR_NO_NAMED_RANGES"
ERR_OUTPUT_DIR = "ERR_OUTPUT_DIR"
ERR_OUTPUT_PATH = "ERR_OUTPUT_PATH"
ERR_PERMISSION_DENIED = "ERR_PERMISSION_DENIED"
ERR_RECORD_COLUMN_TYPE = "ERR_RECORD_COLUMN_TYPE"
ERR_RECORD_DATA_TYPE = "ERR_RECORD_DATA_TYPE"
ERR_RECORD_KEY_TYPE = "ERR_RECORD_KEY_TYPE"
ERR_RECORD_LENGTH_MISMATCH = "ERR_RECORD_LENGTH_MISMATCH"
ERR_RUN_MAIN_FAILED = "ERR_RUN_MAIN_FAILED"
ERR_SEARCH_PATH = "ERR_SEARCH_PATH"
ERR_TEMPLATE_NAME = "ERR_TEMPLATE_NAME"
ERR_TEMPLATE_NOT_FOUND = "ERR_TEMPLATE_NOT_FOUND"
ERR_TEMPLATE_TYPE = "ERR_TEMPLATE_TYPE"
ERR_TEMPLATE_WRITE_FAILED = "ERR_TEMPLATE_WRITE_FAILED"


def ascii_safe(value: Any) -> str:
    """将任意值转换为仅含 ASCII 字符的字符串。"""
    return str(value).encode("ascii", "backslashreplace").decode("ascii")


def format_user_error(code: str, english: str, chinese: str) -> str:
    """生成 Run main 边界可读且只含 ASCII 的双语错误信息。

    Args:
        code: ASCII 错误码，如 ERR_MISSING_NAMED_RANGE。
        english: 英文错误信息。若含非 ASCII 字符，会被转义。
        chinese: 中文补充说明，会被转义为 ASCII-safe 形式。

    Returns:
        以 ASCII 错误码和英文信息开头，并附带转义中文说明的错误文本。
    """
    return f"{code}: {ascii_safe(english)} | CN_ESCAPED: {ascii_safe(chinese)}"
