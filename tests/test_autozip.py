from zipfile import ZipFile

import pytest

import sheetflow.autozip as autozip_module
from sheetflow.autozip import _extract_year_from_filename, _zip_file


def test_zip_file_moves_archive_to_year_directory(tmp_path):
    source = tmp_path / "2024_05.xls"
    source.write_text("tax data")

    assert _zip_file(source) is True

    target = tmp_path / "2024" / "2024_05.zip"
    assert target.exists()
    assert not source.exists()

    with ZipFile(target) as archive:
        assert archive.namelist() == ["2024_05.xls"]
        assert archive.read("2024_05.xls") == b"tax data"


def test_zip_file_replaces_existing_year_archive(tmp_path):
    source = tmp_path / "2024_05.xls"
    source.write_text("new data")
    year_dir = tmp_path / "2024"
    year_dir.mkdir()
    target = year_dir / "2024_05.zip"
    target.write_text("old archive")

    assert _zip_file(source) is True

    with ZipFile(target) as archive:
        assert archive.read("2024_05.xls") == b"new data"


def test_extract_year_from_filename_returns_none_for_unmatched_name(tmp_path):
    assert _extract_year_from_filename(tmp_path / "report.docx", "%Y_%m") is None


def test_auto_zip_raises_when_any_file_fails(monkeypatch, tmp_path):
    source = tmp_path / "2024_05.xls"
    source.write_text("tax data")

    monkeypatch.setattr(autozip_module, "_zip_file", lambda *args: False)

    @autozip_module.auto_zip()
    def process_files():
        return tmp_path

    with pytest.raises(RuntimeError, match="ERR_ARCHIVE_PARTIAL"):
        process_files()

    assert source.exists()


def test_auto_zip_rejects_colliding_archive_names_before_processing(tmp_path):
    first = tmp_path / "Report.xls"
    second = tmp_path / "report.xlsx"
    first.write_text("xls data")
    second.write_text("xlsx data")

    @autozip_module.auto_zip(max_workers=2)
    def process_files():
        return tmp_path

    with pytest.raises(RuntimeError, match="ERR_ARCHIVE_PARTIAL"):
        process_files()

    assert first.read_text() == "xls data"
    assert second.read_text() == "xlsx data"
    assert not (tmp_path / "Report.zip").exists()
    assert not (tmp_path / "report.zip").exists()
