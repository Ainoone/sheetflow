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


def test_zip_file_missing_source_leaves_no_archive(tmp_path):
    assert _zip_file(tmp_path / "missing.xls") is False
    assert list(tmp_path.iterdir()) == []


@pytest.mark.parametrize("error_type", [FileNotFoundError, PermissionError, OSError])
def test_zip_file_write_failure_preserves_source_and_existing_archive(
    monkeypatch, tmp_path, error_type,
):
    source = tmp_path / "report.xls"
    source.write_bytes(b"source data")
    archive = tmp_path / "report.zip"
    archive.write_bytes(b"existing archive")

    def fail_write(self, *args, **kwargs):
        self.writestr("partial", b"incomplete data")
        raise error_type("source read failed")

    monkeypatch.setattr(ZipFile, "write", fail_write)

    assert _zip_file(source) is False
    assert source.read_bytes() == b"source data"
    assert archive.read_bytes() == b"existing archive"
    assert set(tmp_path.iterdir()) == {source, archive}


def test_zip_file_keeps_completed_archive_when_source_delete_fails(monkeypatch, tmp_path):
    source = tmp_path / "2024_05.xls"
    source.write_bytes(b"source data")
    original_unlink = type(source).unlink

    def fail_source_unlink(path, *args, **kwargs):
        if path == source:
            raise PermissionError("source locked")
        return original_unlink(path, *args, **kwargs)

    monkeypatch.setattr(type(source), "unlink", fail_source_unlink)

    assert _zip_file(source) is False
    assert source.read_bytes() == b"source data"
    with ZipFile(tmp_path / "2024" / "2024_05.zip") as archive:
        assert archive.read(source.name) == b"source data"


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
