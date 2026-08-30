import pytest

from sheetflow.common import ascii_safe
from sheetflow import pipeline
from sheetflow.pipeline import _guard_named_ranges


class FakeRange:
    def __init__(self):
        self.value = None


class FakeSheet:
    def __init__(self, name):
        self.name = name
        self.names = []
        self.book = FakeBook()
        self.values = {}

    def range(self, address):
        fake_range = FakeRange()
        self.values[address] = fake_range
        return fake_range


class FakeName:
    def __init__(self, name, value):
        self.name = name
        self.refers_to_range = FakeRefersToRange(value)


class FakeRefersToRange:
    def __init__(self, value):
        self.value = value

    def options(self, **kwargs):
        return self


class FakeSheets:
    def __init__(self):
        self.added = []

    def add(self, name):
        sheet = FakeSheet(name)
        self.added.append(sheet)
        return sheet


class FakeBook:
    def __init__(self):
        self.sheets = FakeSheets()


def test_guard_named_ranges_raises_formatted_error_for_empty_sheet():
    sheet = FakeSheet("资产负债表")

    with pytest.raises(RuntimeError) as exc_info:
        _guard_named_ranges(sheet)

    message = str(exc_info.value)
    assert message.startswith("ERR_NO_NAMED_RANGES:")
    assert "No named ranges on active sheet" in message
    assert "sheet=\\u8d44\\u4ea7\\u8d1f\\u503a\\u8868" in message
    assert ascii_safe("工作表「资产负债表」没有命名区域。") in message
    assert message == ascii_safe(message)


def test_process_sheet_raises_formatted_error_for_mixed_record_shape(monkeypatch, tmp_path):
    sheet = FakeSheet("Sheet1")
    sheet.names = [
        FakeName("Sheet1!Template", "合同"),
        FakeName("Sheet1!CN", ["公司A", "公司B"]),
        FakeName("Sheet1!City", "上海"),
    ]

    monkeypatch.setattr(pipeline, "require_template", lambda template: tmp_path / "合同.docx")

    with pytest.raises(RuntimeError) as exc_info:
        pipeline.process_sheet(sheet, tmp_path, filename=lambda record: f"{record['CN']}.docx")

    message = str(exc_info.value)
    assert message.startswith("ERR_MIXED_RECORD_SHAPE:")
    assert "list_fields=CN" in message
    assert "scalar_fields=City" in message
    assert ascii_safe("工作表「Sheet1」命名区域数据混合了列表字段和标量字段。") in message
    assert message == ascii_safe(message)


def test_process_sheet_keeps_partial_rows_and_skips_only_fully_empty_rows(
    monkeypatch,
    tmp_path,
):
    sheet = FakeSheet("Sheet1")
    sheet.names = [
        FakeName("Sheet1!Template", "合同"),
        FakeName("Sheet1!CN", ["公司A", "公司B", None]),
        FakeName("Sheet1!City", ["上海", None, None]),
    ]
    captured = {}

    monkeypatch.setattr(pipeline, "require_template", lambda template: tmp_path / "合同.docx")

    def capture_render(records, template, output_dir, *, filename):
        captured["records"] = list(records)

    monkeypatch.setattr(pipeline, "render_docx", capture_render)

    pipeline.process_sheet(
        sheet,
        tmp_path,
        filename=lambda record: f"{record['CN']}.docx",
    )

    assert captured["records"] == [
        {"CN": "公司A", "City": "上海"},
        {"CN": "公司B", "City": None},
    ]
