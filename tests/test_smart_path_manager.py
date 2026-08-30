import sys
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pytest

from sheetflow.common import ascii_safe
from sheetflow.smart_path_manager import SmartPathManager


def test_resolve_target_treats_existing_hour_file_as_conflict(tmp_path: Path):
    now = datetime(2026, 5, 12, 14, 30)
    (tmp_path / "26_05_12_14").write_text("not a directory")
    manager = SmartPathManager(tmp_path)

    assert manager._resolve_target_under(tmp_path, now) == tmp_path / "26_05_12_1430"


def test_resolve_target_treats_existing_minute_file_as_conflict(tmp_path: Path):
    now = datetime(2026, 5, 12, 14, 30)
    (tmp_path / "26_05_12_14").mkdir()
    (tmp_path / "26_05_12_1430").write_text("not a directory")
    manager = SmartPathManager(tmp_path)

    assert manager._resolve_target_under(tmp_path, now) == tmp_path / "26_05_12_1430_01"


def test_invalid_path_error_is_ascii_only():
    with pytest.raises(ValueError) as exc_info:
        SmartPathManager("不存在")

    message = str(exc_info.value)
    assert message.startswith("ERR_INVALID_PATH:")
    assert "CN_ESCAPED:" in message
    assert message == ascii_safe(message)


def test_desktop_path_uses_windows_redirected_location(monkeypatch, tmp_path: Path):
    redirected = tmp_path / "OneDrive" / "Desktop"
    redirected.mkdir(parents=True)

    class FakeKey:
        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, traceback):
            return False

    fake_winreg = SimpleNamespace(
        HKEY_CURRENT_USER=object(),
        OpenKey=lambda *args: FakeKey(),
        QueryValueEx=lambda key, name: (str(redirected), 0),
    )
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setitem(sys.modules, "winreg", fake_winreg)

    assert SmartPathManager._desktop_path() == redirected


def test_missing_desktop_falls_back_to_managed_directory(monkeypatch, tmp_path: Path):
    now = datetime(2026, 5, 12, 14, 30)
    (tmp_path / "existing.txt").write_text("data")
    manager = SmartPathManager(tmp_path)
    monkeypatch.setattr(manager, "_desktop_path", lambda: None)

    created = manager.mkdir_if_needed(now=now, threshold=1)

    assert created == tmp_path / "26_05_12_14"
    assert created.is_dir()
