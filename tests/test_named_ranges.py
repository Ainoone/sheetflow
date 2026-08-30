import datetime

import pytest

from sheetflow.common import ascii_safe
from sheetflow.named_ranges import (
    NamedRangeDict,
    build_named_range_map,
    create_named_ranges_from_dict,
    create_sheet_named_ranges,
)


class FakeBook:
    def __init__(self):
        self.names = [FakeName("BookGlobal", "global")]


class FakeName:
    def __init__(self, name, value=None, refers_to="='Sheet1'!$A$1"):
        self.name = name
        self.refers_to_range = FakeRefersToRange(value)
        self.refers_to = refers_to
        self.deleted = False

    def delete(self):
        self.deleted = True


class FakeRefersToRange:
    def __init__(self, value):
        self.value = value
        self.options_kwargs = None

    def options(self, **kwargs):
        self.options_kwargs = kwargs
        return self


class FakeNames(list):
    def __init__(self, values=(), *, should_fail=False):
        super().__init__(values)
        self.should_fail = should_fail

    def add(self, name, refers_to):
        if self.should_fail:
            raise RuntimeError("restore failed")
        restored = FakeName(name, refers_to=refers_to)
        self.append(restored)
        return restored


class FakeTargetRange:
    def __init__(self, address, *, should_fail=False):
        self.address = address
        self.should_fail = should_fail
        self.assigned_name = None

    @property
    def name(self):
        return self.assigned_name

    @name.setter
    def name(self, value):
        if self.should_fail:
            raise RuntimeError("invalid address")
        self.assigned_name = value


class FakeSheet:
    def __init__(
        self,
        name="Sheet1",
        names=None,
        failing_addresses=None,
        restore_should_fail=False,
    ):
        self.name = name
        self.names = FakeNames(names or [], should_fail=restore_should_fail)
        self.book = FakeBook()
        self.failing_addresses = set(failing_addresses or ())
        self.ranges = {}

    def range(self, address):
        target = FakeTargetRange(
            address,
            should_fail=address in self.failing_addresses,
        )
        self.ranges[address] = target
        return target


def test_named_range_dict_cleans_sheet_prefix_and_reads_sheet_names_only():
    local_name = FakeName("'资产负债表'!Template", "报表分析")
    sheet = FakeSheet("资产负债表", names=[local_name])

    data = NamedRangeDict(sheet).data

    assert data == {"Template": "报表分析"}
    assert "BookGlobal" not in data
    assert local_name.refers_to_range.options_kwargs == {
        "dates": datetime.date,
        "empty": None,
    }


def test_named_range_dict_handles_exclamation_mark_in_sheet_name():
    sheet = FakeSheet(
        "Sales!2026",
        names=[FakeName("'Sales!2026'!CN", "Acme")],
    )

    assert NamedRangeDict(sheet).data == {"CN": "Acme"}


def test_named_range_dict_data_is_lazy_and_cached():
    sheet = FakeSheet("Sheet1", names=[FakeName("Sheet1!CN", "Acme")])
    named_ranges = NamedRangeDict(sheet)

    assert named_ranges.data == {"CN": "Acme"}

    sheet.names.append(FakeName("Sheet1!Template", "个税申报表"))
    assert named_ranges.data == {"CN": "Acme"}


def test_build_named_range_map_filters_none_pairs():
    result = build_named_range_map({
        "Named": ["CN", None, "Template", "Skip"],
        "Range": ["A1", "B1", "C1", None],
    })

    assert result == {"CN": "A1", "Template": "C1"}


def test_build_named_range_map_wraps_single_cell_values():
    result = build_named_range_map({"Named": "CN", "Range": "A1"})

    assert result == {"CN": "A1"}


def test_build_named_range_map_rejects_scalar_list_length_mismatch():
    with pytest.raises(ValueError, match="ERR_NAMED_RANGE_MAP"):
        build_named_range_map({"Named": "CN", "Range": ["A1", "B1"]})


def test_build_named_range_map_raises_formatted_error_for_length_mismatch():
    with pytest.raises(ValueError) as exc_info:
        build_named_range_map({"Named": ["CN", "Template"], "Range": ["A1"]})

    message = str(exc_info.value)
    assert message.startswith("ERR_NAMED_RANGE_MAP:")
    assert "2 names vs 1 ranges" in message
    assert ascii_safe("Named/Range 长度不一致") in message
    assert message == ascii_safe(message)


def test_build_named_range_map_rejects_case_insensitive_duplicates():
    with pytest.raises(ValueError, match="ERR_NAMED_RANGE_MAP"):
        build_named_range_map(
            {"Named": ["CN", "cn"], "Range": ["A1", "B1"]}
        )


def test_create_sheet_named_ranges_overwrites_existing_names():
    existing = FakeName("'Sheet1'!CN")
    sheet = FakeSheet("Sheet1", names=[existing])

    result = create_sheet_named_ranges(sheet, {"CN": "A1", "Template": "B1"})

    assert result == {"ok": ["CN", "Template"], "failed": []}
    assert existing.deleted is True
    assert sheet.ranges["A1"].name == "'Sheet1'!CN"
    assert sheet.ranges["B1"].name == "'Sheet1'!Template"


def test_create_sheet_named_ranges_overwrites_unquoted_sheet_prefix():
    # 真实 xlwings 对简单工作表名返回不带引号前缀（"Sheet1!CN"），与创建时
    # 写入的带引号形式（"'Sheet1'!CN"）不一致。覆盖路径按名称本体匹配，
    # 显式删除旧定义，不能依赖 Excel 的重名赋值静默替换行为。
    existing = FakeName("Sheet1!CN")
    sheet = FakeSheet("Sheet1", names=[existing])

    result = create_sheet_named_ranges(sheet, {"CN": "A1"})

    assert result == {"ok": ["CN"], "failed": []}
    assert existing.deleted is True
    assert sheet.ranges["A1"].name == "'Sheet1'!CN"


def test_create_sheet_named_ranges_matches_escaped_sheet_name_prefix():
    existing = FakeName("'O''Brien'!CN")
    sheet = FakeSheet("O'Brien", names=[existing])

    result = create_sheet_named_ranges(sheet, {"CN": "A1"})

    assert result == {"ok": ["CN"], "failed": []}
    assert existing.deleted is True
    assert sheet.ranges["A1"].name == "'O''Brien'!CN"


def test_create_sheet_named_ranges_handles_exclamation_mark_in_sheet_name():
    existing = FakeName("'Sales!2026'!CN")
    sheet = FakeSheet("Sales!2026", names=[existing])

    result = create_sheet_named_ranges(sheet, {"CN": "A1"})

    assert result == {"ok": ["CN"], "failed": []}
    assert existing.deleted is True
    assert sheet.ranges["A1"].name == "'Sales!2026'!CN"


def test_create_sheet_named_ranges_matches_names_case_insensitively():
    existing = FakeName("Sheet1!cn")
    sheet = FakeSheet("Sheet1", names=[existing])

    result = create_sheet_named_ranges(sheet, {"CN": "A1"})

    assert result == {"ok": ["CN"], "failed": []}
    assert existing.deleted is True
    assert sheet.ranges["A1"].name == "'Sheet1'!CN"


@pytest.mark.parametrize(
    "sheet_name",
    [
        "Sheet1",
        "Sheet 1",
        "扣缴个人所得税申报表",
        "扣缴个人所得税申报表（适用于综合所得预扣预缴）",
        "Tax Report (Prepay)",
        "A-B",
        "A.B",
        "A,B",
        "A;B",
        "A+B",
        "A&B",
        "A#B",
        "A%B",
        "2026",
    ],
)
def test_create_sheet_named_ranges_quotes_sheet_names(sheet_name):
    sheet = FakeSheet(sheet_name)

    result = create_sheet_named_ranges(sheet, {"CN": "A1"})

    assert result == {"ok": ["CN"], "failed": []}
    assert sheet.ranges["A1"].name == f"'{sheet_name}'!CN"


def test_create_sheet_named_ranges_escapes_single_quotes_in_sheet_name():
    sheet = FakeSheet("O'Brien", names=[])

    result = create_sheet_named_ranges(sheet, {"CN": "A1"})

    assert result == {"ok": ["CN"], "failed": []}
    assert sheet.ranges["A1"].name == "'O''Brien'!CN"


def test_create_sheet_named_ranges_collects_failed_names():
    sheet = FakeSheet("Sheet1", failing_addresses={"Z999"})

    result = create_sheet_named_ranges(sheet, {"CN": "A1", "Bad": "Z999"})

    assert result == {"ok": ["CN"], "failed": ["Bad"]}
    assert sheet.ranges["A1"].name == "'Sheet1'!CN"


def test_create_sheet_named_ranges_restores_existing_name_after_failure():
    existing = FakeName("'Sheet1'!CN", refers_to="='Sheet1'!$C$3")
    sheet = FakeSheet("Sheet1", names=[existing], failing_addresses={"A1"})

    result = create_sheet_named_ranges(sheet, {"CN": "A1"})

    assert result == {"ok": [], "failed": ["CN"]}
    restored = [item for item in sheet.names if not item.deleted]
    assert len(restored) == 1
    # Worksheet.Names.Add 的 name 参数必须是裸名。
    assert restored[0].name == "CN"
    assert restored[0].refers_to == "='Sheet1'!$C$3"


def test_create_sheet_named_ranges_raises_when_restore_fails():
    existing = FakeName("'Sheet1'!CN", refers_to="='Sheet1'!$C$3")
    sheet = FakeSheet(
        "Sheet1",
        names=[existing],
        failing_addresses={"A1"},
        restore_should_fail=True,
    )

    with pytest.raises(RuntimeError, match="ERR_NAMED_RANGE_RESTORE_FAILED"):
        create_sheet_named_ranges(sheet, {"CN": "A1"})


def test_create_named_ranges_from_dict_builds_and_creates_names():
    sheet = FakeSheet("Sheet1")

    result = create_named_ranges_from_dict(
        sheet,
        {"Named": ["CN", "Template"], "Range": ["A1", "B1"]},
    )

    assert result == {"ok": ["CN", "Template"], "failed": []}
    assert sheet.ranges["A1"].name == "'Sheet1'!CN"
    assert sheet.ranges["B1"].name == "'Sheet1'!Template"
