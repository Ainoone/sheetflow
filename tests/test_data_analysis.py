import importlib.util
from pathlib import Path


SCRIPT = Path(__file__).parents[1] / "py_script" / "数据分析.py"
SPEC = importlib.util.spec_from_file_location("data_analysis_script", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


class FakeRange:
    def __init__(self, should_fail=False):
        self.should_fail = should_fail
        self.value = None

    def set_value(self, value):
        if self.should_fail:
            raise RuntimeError("invalid target")
        self.value = value


class FakeSheet:
    def __init__(self):
        self.ranges = {"A1": FakeRange(), "BAD": FakeRange(should_fail=True)}

    def range(self, address):
        target = self.ranges[address]

        class ValueProxy:
            @property
            def value(self):
                return target.value

            @value.setter
            def value(self, value):
                target.set_value(value)

        return ValueProxy()


def test_write_fields_reports_failures_without_hiding_them():
    failed = MODULE._write_fields(FakeSheet(), {"A1": "ok", "BAD": "bad"})

    assert [key for key, _ in failed] == ["BAD"]


def test_write_fields_successfully_writes_all_fields():
    sheet = FakeSheet()

    failed = MODULE._write_fields(sheet, {"A1": "ok"})

    assert failed == []
    assert sheet.ranges["A1"].value == "ok"
