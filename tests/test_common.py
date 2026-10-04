from pathlib import Path

import pytest

from sheetflow import common
from sheetflow.common import (
    ascii_safe,
    create_output_dir_for_workbook,
    format_user_error,
    pop_required,
    pop_required_many,
    run_main,
)


def assert_ascii_only(value: str) -> None:
    assert value == ascii_safe(value)


class FakeWorkbook:
    def __init__(self, fullname: Path):
        self.fullname = str(fullname)


def test_pop_required_returns_and_removes_value():
    data = {"Template": "个税申报表", "CN": "Acme"}

    assert pop_required(data, "Template") == "个税申报表"
    assert data == {"CN": "Acme"}


def test_pop_required_raises_runtime_error_with_sheet_name():
    with pytest.raises(RuntimeError) as exc_info:
        pop_required({}, "Template", source="Sheet1")

    message = str(exc_info.value)
    assert message.startswith("ERR_MISSING_NAMED_RANGE:")
    assert "Missing named range(s): Template" in message
    assert "sheet=Sheet1" in message
    assert "CN_ESCAPED: " in message
    assert ascii_safe("工作表「Sheet1」缺少 Template 命名区域。") in message
    assert_ascii_only(message)


def test_pop_required_raises_with_ascii_safe_chinese_sheet_name():
    with pytest.raises(RuntimeError) as exc_info:
        pop_required({}, "Template", source="资产负债表")

    message = str(exc_info.value)
    assert "ERR_MISSING_NAMED_RANGE" in message
    assert "Template" in message
    assert f"sheet={ascii_safe('资产负债表')}" in message
    assert ascii_safe("工作表「资产负债表」缺少 Template 命名区域。") in message
    assert_ascii_only(message)


def test_pop_required_many_does_not_mutate_when_any_key_is_missing():
    data = {"start": "2024-01-01", "end": "2024-01-31", "CN": "Acme"}

    with pytest.raises(RuntimeError) as exc_info:
        pop_required_many(data, ("start", "end", "freq"))

    message = str(exc_info.value)
    assert "ERR_MISSING_NAMED_RANGE" in message
    assert "freq" in message
    assert ascii_safe("缺少 freq 命名区域。") in message
    assert_ascii_only(message)
    assert data == {"start": "2024-01-01", "end": "2024-01-31", "CN": "Acme"}


def test_pop_required_many_returns_values_in_requested_keys():
    data = {"start": 1, "end": 2, "freq": "M", "CN": "Acme"}

    assert pop_required_many(data, ("start", "end", "freq")) == {
        "start": 1,
        "end": 2,
        "freq": "M",
    }
    assert data == {"CN": "Acme"}


def test_create_output_dir_for_workbook_creates_under_workbook_parent(tmp_path: Path):
    workbook = FakeWorkbook(tmp_path / "book.xlsx")

    result = create_output_dir_for_workbook(workbook)

    assert result.parent == tmp_path
    assert result.is_dir()


def test_create_output_dir_for_workbook_raises_when_manager_returns_none(monkeypatch, tmp_path):
    class FakeManager:
        def __init__(self, path):
            self.path = path

        def mkdir_if_needed(self):
            return None

    monkeypatch.setattr(common, "SmartPathManager", FakeManager)

    with pytest.raises(RuntimeError) as exc_info:
        create_output_dir_for_workbook(FakeWorkbook(tmp_path / "book.xlsx"))

    message = str(exc_info.value)
    assert message.startswith("ERR_OUTPUT_DIR:")
    assert "Unable to create output directory" in message
    assert ascii_safe("无法创建输出目录") in message
    assert_ascii_only(message)


def test_ascii_safe_escapes_non_ascii_values():
    assert ascii_safe("资产负债表") == "\\u8d44\\u4ea7\\u8d1f\\u503a\\u8868"


def test_format_user_error_keeps_ascii_prefix_and_cn_escaped_suffix():
    message = format_user_error(
        "ERR_TEST",
        "Sheet name: 资产负债表",
        "工作表名称：资产负债表",
    )

    assert message.startswith("ERR_TEST: Sheet name: \\u8d44")
    assert message.endswith(f" | CN_ESCAPED: {ascii_safe('工作表名称：资产负债表')}")
    assert_ascii_only(message)


def test_run_main_wraps_unformatted_exception_without_context_display():
    @run_main("data_analysis")
    def fail():
        raise ValueError("bad sheet 资产负债表")

    with pytest.raises(RuntimeError) as exc_info:
        fail()

    message = str(exc_info.value)
    assert message.startswith("ERR_RUN_MAIN_FAILED:")
    assert "Run main failed: data_analysis" in message
    assert "bad sheet \\u8d44\\u4ea7\\u8d1f\\u503a\\u8868" in message
    assert ascii_safe("data_analysis 运行失败：bad sheet 资产负债表") in message
    assert_ascii_only(message)
    assert exc_info.value.__suppress_context__ is True


def test_run_main_does_not_wrap_formatted_error():
    @run_main("data_analysis")
    def fail():
        raise RuntimeError("ERR_ALREADY_FORMATTED: details | CN_ESCAPED: \\u8be6\\u60c5")

    with pytest.raises(RuntimeError) as exc_info:
        fail()

    assert str(exc_info.value) == "ERR_ALREADY_FORMATTED: details | CN_ESCAPED: \\u8be6\\u60c5"
