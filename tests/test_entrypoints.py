import runpy
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from sheetflow import personal_income_tax


def _load_script(name: str):
    return runpy.run_path(Path(__file__).parents[1] / "py_script" / name)


def test_namedrange_entrypoint_rejects_partial_creation():
    module = _load_script("namedrange.py")

    with pytest.raises(RuntimeError, match="ERR_NAMED_RANGE_PARTIAL"):
        module["_raise_if_failed"]({"ok": ["CN"], "failed": ["Bad"]})


def test_namedrange_entrypoint_requires_named_and_range(monkeypatch):
    module = _load_script("namedrange.py")
    main = module["main"]
    main_globals = main.__wrapped__.__globals__
    active_sheet = SimpleNamespace(name="named range")
    workbook = SimpleNamespace(
        sheets=SimpleNamespace(active=active_sheet),
        save=Mock(),
    )

    monkeypatch.setattr(main_globals["xw"].Book, "caller", lambda: workbook)
    monkeypatch.setitem(
        main_globals,
        "NamedRangeDict",
        lambda sheet: SimpleNamespace(data={"Sheet_Name": "Target"}),
    )

    with pytest.raises(RuntimeError, match="ERR_MISSING_NAMED_RANGE"):
        main()

    workbook.save.assert_not_called()


def test_tax_entrypoint_validates_period_before_output_side_effects(monkeypatch):
    module = _load_script("税务.py")
    main = module["main"].__wrapped__
    main_globals = main.__globals__

    class CallerSheet:
        name = "Input"

    class CallerBook:
        sheets = SimpleNamespace(active=CallerSheet())

    create_output = Mock()
    generate_tax = Mock()
    monkeypatch.setattr(main_globals["xw"].Book, "caller", lambda: CallerBook())
    monkeypatch.setitem(
        main_globals,
        "NamedRangeDict",
        lambda sheet: SimpleNamespace(
            data={
                "Template": "个税压缩包",
                "start": date(2024, 1, 1),
                "end": date(2024, 1, 31),
                "freq": "X",
            }
        ),
    )
    monkeypatch.setitem(main_globals, "create_output_dir_for_workbook", create_output)
    monkeypatch.setitem(main_globals, "generate_personal_income_tax", generate_tax)

    with pytest.raises(ValueError, match="Invalid freq"):
        main()

    create_output.assert_not_called()
    generate_tax.assert_not_called()


def test_business_entrypoint_defers_output_dir_until_sheet_validation(monkeypatch):
    module = _load_script("工商.py")
    main_globals = module["main"].__wrapped__.__globals__
    active_sheet = SimpleNamespace(name="Input")
    workbook = SimpleNamespace(sheets=SimpleNamespace(active=active_sheet))
    create_output = Mock(return_value=Path("unused"))

    monkeypatch.setattr(main_globals["xw"].Book, "caller", lambda: workbook)
    monkeypatch.setitem(main_globals, "create_output_dir_for_workbook", create_output)

    def reject_sheet(sheet, output_dir, *, filename):
        assert callable(output_dir)
        raise FileNotFoundError("template missing")

    monkeypatch.setitem(main_globals, "process_sheet", reject_sheet)

    with pytest.raises(FileNotFoundError, match="template missing"):
        module["main"].__wrapped__()

    create_output.assert_not_called()


def test_tax_entrypoint_resolves_word_template_before_output_side_effects(monkeypatch):
    module = _load_script("税务.py")
    main_globals = module["main"].__wrapped__.__globals__

    class CallerSheet:
        name = "Input"

    class CallerBook:
        sheets = SimpleNamespace(active=CallerSheet())

    create_output = Mock()
    require_template = Mock(side_effect=FileNotFoundError("template missing"))
    monkeypatch.setattr(main_globals["xw"].Book, "caller", lambda: CallerBook())
    monkeypatch.setitem(
        main_globals,
        "NamedRangeDict",
        lambda sheet: SimpleNamespace(
            data={
                "Template": "企业所得税",
                "start": date(2024, 1, 1),
                "end": date(2024, 1, 31),
                "freq": "M",
                "CN": "公司A",
            }
        ),
    )
    monkeypatch.setitem(main_globals, "create_output_dir_for_workbook", create_output)
    monkeypatch.setitem(main_globals, "require_template", require_template)

    with pytest.raises(FileNotFoundError, match="template missing"):
        module["main"].__wrapped__()

    require_template.assert_called_once_with("企业所得税")
    create_output.assert_not_called()


def test_tax_entrypoint_rejects_invalid_filename_before_creating_directory(
    monkeypatch, tmp_path,
):
    module = _load_script("税务.py")
    main_globals = module["main"].__wrapped__.__globals__
    workbook = SimpleNamespace(sheets=SimpleNamespace(active=SimpleNamespace(name="Input")))
    create_output = Mock(return_value=tmp_path)
    monkeypatch.setattr(main_globals["xw"].Book, "caller", lambda: workbook)
    monkeypatch.setitem(
        main_globals, "NamedRangeDict",
        lambda sheet: SimpleNamespace(data={
            "Template": "企业所得税", "CN": "A/B公司",
            "start": date(2024, 1, 1), "end": date(2024, 2, 29), "freq": "M",
        }),
    )
    monkeypatch.setitem(
        main_globals, "require_template", lambda *args: tmp_path / "template.docx",
    )
    monkeypatch.setitem(main_globals, "create_output_dir_for_workbook", create_output)

    with pytest.raises(ValueError, match="ERR_OUTPUT_PATH") as exc_info:
        module["main"]()

    assert str(exc_info.value).isascii()
    create_output.assert_not_called()


class FakeRange:
    def __init__(self, should_fail=False):
        self.should_fail = should_fail

    @property
    def value(self):
        return None

    @value.setter
    def value(self, value):
        if self.should_fail:
            raise RuntimeError("template write failed")


class FakeSheet:
    def __init__(self, should_fail=False):
        self.should_fail = should_fail

    def range(self, address):
        return FakeRange(self.should_fail)


class FakeWorkbook:
    def __init__(self, should_fail=False):
        self.sheets = [FakeSheet(should_fail)]

    def save(self, path):
        Path(path).write_text("generated")

    def close(self):
        pass


class FakeBooks:
    def __init__(self):
        self.open_count = 0

    def open(self, path):
        self.open_count += 1
        return FakeWorkbook(should_fail=self.open_count == 2)


class FakeApp:
    def __init__(self):
        self.books = FakeBooks()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


def test_personal_income_tax_cleans_partial_outputs(monkeypatch, tmp_path):
    app = FakeApp()
    monkeypatch.setattr(personal_income_tax.xw, "App", lambda **kwargs: app)
    monkeypatch.setattr(
        personal_income_tax,
        "search_template_file_cached",
        lambda *args, **kwargs: tmp_path / "2018.xls",
    )

    with pytest.raises(RuntimeError, match="template write failed"):
        personal_income_tax.generate_personal_income_tax(
            {},
            [
                (date(2018, 1, 1), date(2018, 1, 31)),
                (date(2018, 2, 1), date(2018, 2, 28)),
            ],
            tmp_path,
        )

    assert not (tmp_path / "2018_01.xls").exists()
    assert not (tmp_path / "2018_02.xls").exists()


def test_personal_income_tax_keeps_preexisting_output_on_failure(monkeypatch, tmp_path):
    preexisting = tmp_path / "2018_02.xls"
    preexisting.write_text("existing tax data")
    app = FakeApp()
    monkeypatch.setattr(personal_income_tax.xw, "App", lambda **kwargs: app)
    monkeypatch.setattr(
        personal_income_tax,
        "search_template_file_cached",
        lambda *args, **kwargs: tmp_path / "2018.xls",
    )

    with pytest.raises(RuntimeError, match="template write failed"):
        personal_income_tax.generate_personal_income_tax(
            {},
            [
                (date(2018, 1, 1), date(2018, 1, 31)),
                (date(2018, 2, 1), date(2018, 2, 28)),
            ],
            tmp_path,
        )

    assert not (tmp_path / "2018_01.xls").exists()
    assert preexisting.read_text() == "existing tax data"


def test_data_analysis_main_does_not_save_after_write_failure(monkeypatch, tmp_path):
    module = _load_script("数据分析.py")
    main_globals = module["main"].__wrapped__.__globals__

    class CallerSheet:
        name = "Input"

    class CallerBook:
        fullname = str(tmp_path / "input.xlsx")
        sheets = SimpleNamespace(active=CallerSheet())

    class FieldSheet:
        def range(self, address):
            return FakeRange(should_fail=address == "BAD")

    class OutputWorkbook(FakeWorkbook):
        def __init__(self):
            self.sheets = [FieldSheet()]
            self.saved = False

        def save(self, path):
            self.saved = True

    output_workbook = OutputWorkbook()

    class OutputApp(FakeApp):
        def __init__(self):
            self.display_alerts = True
            self.screen_updating = True
            self.books = SimpleNamespace(open=lambda path: output_workbook)

    monkeypatch.setattr(main_globals["xw"].Book, "caller", lambda: CallerBook())
    monkeypatch.setitem(
        main_globals,
        "NamedRangeDict",
        lambda sheet: SimpleNamespace(
            data={"Template": "报表分析", "A1": "ok", "BAD": "bad"}
        ),
    )
    monkeypatch.setitem(
        main_globals,
        "require_template",
        lambda *args, **kwargs: tmp_path / "报表分析.xlsx",
    )
    monkeypatch.setitem(
        main_globals,
        "create_output_dir_for_workbook",
        lambda workbook: tmp_path,
    )
    monkeypatch.setattr(main_globals["xw"], "App", lambda **kwargs: OutputApp())

    with pytest.raises(RuntimeError, match="ERR_TEMPLATE_WRITE_FAILED"):
        module["main"]()

    assert output_workbook.saved is False
