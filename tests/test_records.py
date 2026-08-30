from types import MappingProxyType

import pytest

from sheetflow.records import RecordSet
from sheetflow.user_errors import ascii_safe


def test_record_set_transposes_columns_to_rows():
    records = RecordSet({"name": ["Alice", "Bob"], "age": [25, 30]})

    assert list(records) == [
        {"name": "Alice", "age": 25},
        {"name": "Bob", "age": 30},
    ]


def test_record_set_dropna_skips_rows_with_none():
    records = RecordSet({"name": ["Alice", None], "age": [25, 30]}, dropna=True)

    assert list(records) == [{"name": "Alice", "age": 25}]


def test_record_set_dropna_checks_none_by_identity():
    class Value:
        def __eq__(self, other):
            raise AssertionError("dropna must not compare cell values")

    value = Value()

    assert list(RecordSet({"value": [value]}, dropna=True)) == [{"value": value}]


def test_record_set_rejects_columns_with_different_lengths():
    with pytest.raises(ValueError) as exc_info:
        RecordSet({"name": ["Alice", "Bob"], "age": [25]})

    message = str(exc_info.value)
    assert message.startswith("ERR_RECORD_LENGTH_MISMATCH:")
    assert "lengths={'name': 2, 'age': 1}" in message
    assert message == ascii_safe(message)


def test_record_set_accepts_empty_mapping():
    records = RecordSet({})

    assert list(records) == []
    assert records.records == []
    assert repr(records) == "RecordSet(keys=[])"


def test_record_set_accepts_read_only_mapping():
    data = MappingProxyType({"name": ["Alice"], "age": [25]})

    assert list(RecordSet(data)) == [{"name": "Alice", "age": 25}]


def test_record_set_snapshots_mapping_and_columns_at_construction():
    data = {"name": ["Alice"], "age": [25]}
    records = RecordSet(data)

    data["name"][0] = "Bob"
    data["age"].append(30)
    del data["name"]
    data["city"] = ["Shanghai", "Beijing"]

    assert list(records) == [{"name": "Alice", "age": 25}]


def test_record_set_repeated_iteration_returns_fresh_record_dicts():
    records = RecordSet({"name": ["Alice"], "age": [25]})

    first = list(records)
    second = list(records)

    assert first == second
    assert first[0] is not second[0]


def test_records_property_returns_an_independent_projection():
    records = RecordSet({"name": ["Alice"], "age": [25]})

    first = records.records
    first[0]["name"] = "Bob"
    first.append({"name": "Carol", "age": 30})

    assert records.records == [{"name": "Alice", "age": 25}]
    assert records.records is not records.records


@pytest.mark.parametrize("data", [None, [], 0])
def test_record_set_rejects_non_mapping_data(data):
    with pytest.raises(TypeError, match="^ERR_RECORD_DATA_TYPE:") as exc_info:
        RecordSet(data)

    assert str(exc_info.value) == ascii_safe(str(exc_info.value))


def test_record_set_rejects_non_string_keys():
    with pytest.raises(TypeError, match="^ERR_RECORD_KEY_TYPE:") as exc_info:
        RecordSet({1: ["Alice"]})

    assert str(exc_info.value) == ascii_safe(str(exc_info.value))


@pytest.mark.parametrize("column", ["AB", ("A", "B"), 1])
def test_record_set_rejects_non_list_columns(column):
    with pytest.raises(TypeError, match="^ERR_RECORD_COLUMN_TYPE:") as exc_info:
        RecordSet({"name": column})

    assert str(exc_info.value) == ascii_safe(str(exc_info.value))


def test_record_set_repr_preserves_key_order():
    records = RecordSet({"name": ["Alice"], "age": [25]})

    assert repr(records) == "RecordSet(keys=['name', 'age'])"
