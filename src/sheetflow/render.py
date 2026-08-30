"""Word 文档批量渲染模块。

渲染前会物化记录并预检输出文件名与路径，避免路径错误留下部分文档。

典型用例：
    >>> from pathlib import Path
    >>> output_dir = Path('output')
    >>> output_dir.mkdir(parents=True, exist_ok=True)
    >>> records = [{'CN': '公司A'}, {'CN': '公司B'}]
    >>> render_docx(records, 'template.docx', output_dir,
    ...             filename=lambda m: f"{m['CN']}.docx")
"""

import contextlib
import logging
from pathlib import Path
from typing import Any, Callable, Generator, Iterable

from docxtpl import DocxTemplate
from more_itertools import always_iterable

from .user_errors import (
    ERR_FILE_NOT_FOUND,
    ERR_NOT_A_FILE,
    ERR_OUTPUT_PATH,
    ERR_PERMISSION_DENIED,
    ERR_TEMPLATE_TYPE,
    ascii_safe,
    format_user_error,
)

logger = logging.getLogger(__name__)


def _resolve_output_path(output_dir: Path, filename: str) -> Path:
    """解析输出文件路径，并拒绝逃出指定输出目录的结果。"""
    # resolve 会折叠 ``..`` 并跟随已有符号链接，边界判断必须基于真实路径。
    output_root = output_dir.resolve()
    output_path = (output_root / filename).resolve()

    if output_path == output_root or not output_path.is_relative_to(output_root):
        raise ValueError(
            format_user_error(
                ERR_OUTPUT_PATH,
                f"Generated file escapes output directory; filename={ascii_safe(filename)}",
                f"生成文件路径超出输出目录：{filename}",
            )
        )

    return output_path


@contextlib.contextmanager
def _open_template(path: str | Path) -> Generator[DocxTemplate, None, None]:
    """校验并打开 .docx 模板，返回 DocxTemplate 上下文。

    会把模板类型、文件不存在、路径不是文件和权限问题转换为带 ERR_* 错误码
    的异常信息，便于 Run main 场景排障。
    """
    full_path = Path(path).resolve()

    if full_path.suffix.lower() != ".docx":
        raise ValueError(
            format_user_error(
                ERR_TEMPLATE_TYPE,
                f"Only .docx templates are supported; file={ascii_safe(full_path.name)}",
                f"仅支持 .docx 文件，收到: {full_path.name}",
            )
        )
    if not full_path.exists():
        raise FileNotFoundError(
            format_user_error(
                ERR_FILE_NOT_FOUND,
                f"File does not exist: {ascii_safe(full_path)}",
                f"文件不存在: {full_path}",
            )
        )
    if not full_path.is_file():
        raise ValueError(
            format_user_error(
                ERR_NOT_A_FILE,
                f"Path is not a file: {ascii_safe(full_path)}",
                f"路径不是文件: {full_path}",
            )
        )

    try:
        tpl = DocxTemplate(full_path)
    except PermissionError:
        logger.error(f"无权限读取文件: {full_path}")
        raise RuntimeError(
            format_user_error(
                ERR_PERMISSION_DENIED,
                f"Permission denied while opening template; template={ascii_safe(full_path)}",
                f"无权限读取文件: {full_path}",
            )
        ) from None
    except Exception as e:
        logger.error(f"打开模板失败: {full_path}, 错误: {e}")
        raise

    # 只转换模板构造阶段的读取错误；render/save 异常保留原始原因和目标路径。
    logger.debug(f"已打开模板: {full_path}")
    yield tpl


def render_docx(
    records: dict[str, Any] | Iterable[dict[str, Any]],
    template: str | Path,
    output_dir: Path | Callable[[], Path],
    *,
    filename: Callable[[dict], str],
) -> Path:
    """批量渲染 Word 模板并保存。

    Args:
        records: 单条记录字典，或记录的可迭代对象
        template: 模板文件路径（.docx)
        output_dir: 已创建的输出目录，或无参数的延迟创建函数。传入函数时，所有记录
            和原始文件名会先完成预检，再调用函数创建目录。
        filename: 从记录生成文件名的可调用对象，如 lambda m: f"{m['CN']}.docx"

    Returns:
        实际使用的输出目录

    Raises:
        FileNotFoundError: 模板文件不存在(ERR_FILE_NOT_FOUND)。
        ValueError: 模板路径不是 .docx 文件(ERR_TEMPLATE_TYPE)，路径不是文件
            (ERR_NOT_A_FILE)，或生成路径逃出输出目录/与本批次已有路径冲突
            (ERR_OUTPUT_PATH)。
        RuntimeError: 模板文件无权限读取(ERR_PERMISSION_DENIED)。
        KeyError: filename 所需的键在记录中不存在

    Note:
        - 会先物化记录并生成本批次全部文件名。若 output_dir 是延迟创建函数，文件名
          生成和直接重复错误会在创建目录前暴露；完整路径会在写文件前统一校验。
        - 实际模板渲染和保存仍按顺序执行，并非事务；此阶段后续记录失败时，先前
          已保存的文档会保留。

    Examples:
        >>> from pathlib import Path
        >>> output_dir = Path('output'); output_dir.mkdir(parents=True, exist_ok=True)
        >>> records = [{'CN': '公司A'}, {'CN': '公司B'}]
        >>> render_docx(records, 'template.docx', output_dir,
        ...             filename=lambda m: f"{m['CN']}.docx")
        PosixPath('output')

        >>> # 自定义文件名
        >>> render_docx(records, 'template.docx', output_dir,
        ...             filename=lambda m: f"{m['CN']}_签字材料.docx")
    """
    # 先冻结一次性迭代器，确保文件名与路径预检覆盖完整批次。
    records = tuple(always_iterable(records, base_type=dict))
    generated_names: list[str] = []
    generated_name_keys: set[str] = set()
    for record in records:
        generated_name = filename(record)
        generated_name_key = generated_name.casefold()
        if generated_name_key in generated_name_keys:
            raise ValueError(
                format_user_error(
                    ERR_OUTPUT_PATH,
                    f"Duplicate generated output filename: {ascii_safe(generated_name)}",
                    f"本批次生成了重复的输出文件名：{generated_name}",
                )
            )
        generated_name_keys.add(generated_name_key)
        generated_names.append(generated_name)

    resolved_output_dir = output_dir() if callable(output_dir) else output_dir
    output_keys: set[str] = set()
    prepared: list[tuple[dict[str, Any], Path]] = []
    for record, generated_name in zip(records, generated_names):
        path = _resolve_output_path(resolved_output_dir, generated_name)
        output_key = str(path).casefold()
        if output_key in output_keys:
            raise ValueError(
                format_user_error(
                    ERR_OUTPUT_PATH,
                    f"Duplicate generated output path: {ascii_safe(path)}",
                    f"本批次生成了重复的输出路径：{path}",
                )
            )
        output_keys.add(output_key)
        prepared.append((record, path))

    with _open_template(template) as doc:
        for record, path in prepared:
            doc.render(record)
            doc.save(path)
            logger.info(f"已保存: {path}")

    return resolved_output_dir
