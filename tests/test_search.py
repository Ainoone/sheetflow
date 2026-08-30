from pathlib import Path

import pytest

from sheetflow.common import ascii_safe
from sheetflow.search import require_template
from sheetflow.search import search_template_file
from sheetflow.search import search_template_file_cached


def test_search_template_file_returns_none_when_missing(tmp_path: Path):
    assert search_template_file("missing", search_path=tmp_path) is None


def test_require_template_returns_found_template(tmp_path: Path):
    template = tmp_path / "合同.docx"
    template.write_text("template")

    assert require_template("合同", search_path=tmp_path) == template


def test_require_template_raises_when_missing(tmp_path: Path):
    with pytest.raises(FileNotFoundError) as exc_info:
        require_template("缺失模板", search_path=tmp_path)

    message = str(exc_info.value)
    assert message.startswith("ERR_TEMPLATE_NOT_FOUND:")
    assert "Template not found: \\u7f3a\\u5931\\u6a21\\u677f.docx" in message
    assert ascii_safe("找不到模板文件: 缺失模板.docx") in message
    assert message == ascii_safe(message)


def test_search_template_file_rejects_parent_traversal(tmp_path: Path):
    template_dir = tmp_path / "templates"
    template_dir.mkdir()
    (tmp_path / "outside.docx").write_text("outside template")

    with pytest.raises(ValueError, match="ERR_TEMPLATE_NAME") as exc_info:
        search_template_file("../outside", search_path=template_dir)

    assert str(exc_info.value) == ascii_safe(exc_info.value)


def test_search_template_file_treats_glob_characters_literally(tmp_path: Path):
    (tmp_path / "contract.docx").write_text("template")

    assert search_template_file("*", search_path=tmp_path) is None


def test_search_template_file_cached_uses_bounded_cache():
    assert search_template_file_cached.cache_parameters()["maxsize"] == 128


def test_search_template_file_cached_does_not_cache_missing_result(tmp_path: Path):
    search_template_file_cached.cache_clear()
    assert search_template_file_cached("later", search_path=tmp_path) is None

    template = tmp_path / "later.docx"
    template.write_text("template")

    assert search_template_file_cached("later", search_path=tmp_path) == template


def test_search_template_file_cached_refreshes_deleted_path(tmp_path: Path):
    search_template_file_cached.cache_clear()
    original = tmp_path / "report.docx"
    original.write_text("old")
    assert search_template_file_cached("report", search_path=tmp_path) == original

    original.unlink()
    replacement_dir = tmp_path / "replacement"
    replacement_dir.mkdir()
    replacement = replacement_dir / "report.docx"
    replacement.write_text("new")

    assert search_template_file_cached("report", search_path=tmp_path) == replacement
