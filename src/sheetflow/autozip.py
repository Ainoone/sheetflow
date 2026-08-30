"""自动压缩模块。

提供自动压缩目录直接包含文件的装饰器，支持并发压缩和智能文件组织。

主要功能：
    - 自动压缩目录直接包含的所有非 .zip 文件
    - 并发处理提高效率
    - 根据文件名日期自动组织到年份目录
    - 压缩后删除原文件

典型用例：
    >>> from pathlib import Path
    >>> @auto_zip()
    ... def process_files():
    ...     return Path('output')
    >>>
    >>> result = process_files()
    >>> print(result)
    PosixPath('output')

    >>> # 自定义配置
    >>> @auto_zip(compress_level=9, date_format="%Y%m%d")
    ... def process_files():
    ...     return Path('output')
"""

from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from functools import wraps
import logging
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from .user_errors import (
    ERR_ARCHIVE_PARTIAL,
    ERR_OUTPUT_DIR,
    ascii_safe,
    format_user_error,
)

logger = logging.getLogger(__name__)


def auto_zip(
    compress_level: int = 6,
    date_format: str = "%Y_%m",
    max_workers: int | None = None
) -> Callable[[Callable[..., Path]], Callable[..., Path]]:
    """自动压缩目录中文件的装饰器。

    Args:
        compress_level: 压缩级别 (0-9)，默认 6。9 为最高压缩率
        date_format: 文件名日期格式，默认 "%Y_%m"
        max_workers: 最大线程数，默认 None(自动)

    Returns:
        装饰器函数

    Raises:
        TypeError: 被装饰函数未返回 pathlib.Path(ERR_OUTPUT_DIR)。
        FileNotFoundError: 返回路径不存在(ERR_OUTPUT_DIR)。
        ValueError: 返回路径不是目录(ERR_OUTPUT_DIR)。
        RuntimeError: 压缩目标冲突或任一文件处理失败(ERR_ARCHIVE_PARTIAL)。

    Examples:
        >>> from pathlib import Path
        >>> # 基本用法
        >>> @auto_zip()
        ... def process_files():
        ...     return Path('output')
        >>>
        >>> result = process_files()
        >>> print(result)
        PosixPath('output')

        >>> # 自定义配置
        >>> @auto_zip(compress_level=9, date_format="%Y%m%d", max_workers=4)
        ... def process_files():
        ...     return Path('output')

    Note:
        - 装饰的函数必须返回 Path 对象
        - 会压缩目录直接包含的所有非 .zip 文件，不递归处理子目录
        - 压缩成功后删除原文件
        - 根据文件名日期自动组织到年份目录
        - 主干名称仅大小写不同或扩展名不同的源文件可能对应同一个 .zip；检测到
          这种跨平台冲突时，会在开始压缩前抛出 ERR_ARCHIVE_PARTIAL
        - 返回原函数的返回值 (Path 对象)
        - 任一文件压缩失败时抛出 ERR_ARCHIVE_PARTIAL
        - 批量处理不是事务；发生单文件失败前已成功生成的压缩包及年份目录会保留，
          对应原文件也已删除
    """
    def decorator(func: Callable[..., Path]) -> Callable[..., Path]:
        """将目录压缩后处理附加到目标函数。"""
        @wraps(func)
        def wrapper(*args, **kwargs) -> Path:
            """执行目标函数并压缩其返回目录中的文件。"""
            path = func(*args, **kwargs)

            # 先验证返回契约，避免后续文件系统操作产生部分副作用。
            if not isinstance(path, Path):
                raise TypeError(
                    format_user_error(
                        ERR_OUTPUT_DIR,
                        f"Decorated function must return pathlib.Path; got={ascii_safe(type(path))}",
                        f"函数必须返回 Path 对象，收到: {type(path)}",
                    )
                )

            if not path.exists():
                raise FileNotFoundError(
                    format_user_error(
                        ERR_OUTPUT_DIR,
                        f"Path does not exist: {ascii_safe(path)}",
                        f"路径不存在: {path}",
                    )
                )

            if not path.is_dir():
                raise ValueError(
                    format_user_error(
                        ERR_OUTPUT_DIR,
                        f"Path is not a directory: {ascii_safe(path)}",
                        f"路径不是目录: {path}",
                    )
                )

            # 仅处理顶层非 ZIP 文件，避免递归处理年份目录和重复压缩。
            files_to_zip = [
                f for f in path.iterdir()
                if f.is_file() and f.suffix.lower() != '.zip'
            ]

            if not files_to_zip:
                logger.info(f"目录中没有需要压缩的文件: {path}")
                return path

            archive_sources: dict[str, Path] = {}
            conflicts: set[str] = set()
            # 并发开始前按大小写不敏感的目标名预检，避免产生部分归档。
            for file_path in files_to_zip:
                archive_key = file_path.with_suffix('.zip').name.casefold()
                previous = archive_sources.get(archive_key)
                if previous is None:
                    archive_sources[archive_key] = file_path
                    continue
                conflicts.update((previous.name, file_path.name))

            if conflicts:
                details = ", ".join(sorted(conflicts))
                raise RuntimeError(
                    format_user_error(
                        ERR_ARCHIVE_PARTIAL,
                        f"Archive target conflict: {ascii_safe(details)}",
                        f"多个源文件对应同一个压缩包：{details}",
                    )
                )

            logger.info(f"开始压缩 {len(files_to_zip)} 个文件...")

            # 每个文件独立压缩；等待全部任务完成后统一汇总失败。
            success_count = 0
            error_count = 0

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_file = {
                    executor.submit(
                        _zip_file,
                        file_path,
                        compress_level,
                        date_format
                    ): file_path
                    for file_path in files_to_zip
                }

                for future in as_completed(future_to_file):
                    file_path = future_to_file[future]
                    try:
                        result = future.result()
                        if result:
                            success_count += 1
                        else:
                            error_count += 1
                    except Exception as e:
                        logger.error(f"压缩文件失败 {file_path}: {e}")
                        error_count += 1

            logger.info(
                f"压缩完成: 成功 {success_count}, 失败 {error_count}"
            )

            if error_count:
                raise RuntimeError(
                    format_user_error(
                        ERR_ARCHIVE_PARTIAL,
                        f"Archive failed: success={success_count}, failed={error_count}",
                        f"压缩失败：成功 {success_count} 个，失败 {error_count} 个",
                    )
                )

            return path

        return wrapper
    return decorator


def _zip_file(
    file_path: Path,
    compress_level: int = 6,
    date_format: str = "%Y_%m"
) -> bool:
    """压缩单个文件并组织到年份目录。

    Args:
        file_path: 要压缩的文件路径
        compress_level: 压缩级别 (0-9)，默认 6
        date_format: 文件名日期格式，默认 "%Y_%m"

    Returns:
        True 表示成功,False 表示失败

    Examples:
        >>> from pathlib import Path
        >>> file = Path('output/2024_05.docx')
        >>> _zip_file(file)
        True

        >>> # 自定义压缩级别
        >>> _zip_file(file, compress_level=9)
        True

        >>> # 自定义日期格式
        >>> file = Path('output/20240502.docx')
        >>> _zip_file(file, date_format="%Y%m%d")
        True

    Note:
        - 压缩成功后删除原文件
        - 如果不含扩展名的完整文件名符合 date_format，会移动到对应年份目录
        - 所有操作成功后才删除原文件
        - 如果目标文件已存在会被覆盖
        - 若删除原文件失败，函数返回 False，已生成或移动的压缩包可能保留
    """
    archive = file_path.with_suffix('.zip')

    try:
        # 先写入归档；只有归档和移动完成后才会删除源文件。
        with ZipFile(
            archive,
            'w',
            compression=ZIP_DEFLATED,
            compresslevel=compress_level
        ) as myzip:
            myzip.write(filename=file_path, arcname=file_path.name)

        logger.debug(f"已压缩: {file_path} -> {archive}")

        # 日期型文件名按年份归档，其他文件保留在当前目录。
        year = _extract_year_from_filename(file_path, date_format)
        if year:
            new_directory = file_path.parent / str(year)
            new_directory.mkdir(exist_ok=True)

            target = new_directory / archive.name

            archive.replace(target)
            archive = target  # 后续异常清理必须指向移动后的归档。
            logger.debug(f"已移动到年份目录: {target}")

        # 最后删除源文件，确保前述归档操作均已完成。
        file_path.unlink()
        logger.info(f"处理完成: {file_path.name}")

        return True

    except FileNotFoundError:
        logger.error(f"文件不存在: {file_path}")
        return False
    except PermissionError as e:
        logger.error(f"权限错误: {file_path}, {e}")
        return False
    except Exception as e:
        logger.error(f"压缩文件时发生错误 {file_path}: {e}")
        # 尽力清理当前归档路径；清理失败仍保留本次处理失败的返回值。
        if archive.exists():
            try:
                archive.unlink()
            except Exception:
                pass
        return False


def _extract_year_from_filename(
    file_path: Path,
    date_format: str
) -> int | None:
    """从文件名提取年份。

    Args:
        file_path: 文件路径
        date_format: 日期格式字符串

    Returns:
        年份（整数），解析失败返回 None

    Examples:
        >>> from pathlib import Path
        >>> _extract_year_from_filename(Path('2024_05.docx'), '%Y_%m')
        2024

        >>> _extract_year_from_filename(Path('20240502.docx'), '%Y%m%d')
        2024

        >>> _extract_year_from_filename(Path('report.docx'), '%Y_%m')
        None

    Note:
        - 只解析文件名（不含扩展名）
        - 解析失败时返回 None,不抛出异常
    """
    try:
        return datetime.strptime(file_path.stem, date_format).year
    except ValueError:
        return None
