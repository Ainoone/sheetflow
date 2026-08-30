from pathlib import Path
from unittest.mock import Mock

import pytest

from sheetflow.common import ascii_safe
from sheetflow import render


class FakeDoc:
    def __init__(self):
        self.rendered = []
        self.saved = []

    def render(self, record):
        self.rendered.append(record)

    def save(self, path):
        self.saved.append(path)


class FakeTemplate:
    def __enter__(self):
        self.doc = FakeDoc()
        return self.doc

    def __exit__(self, exc_type, exc, tb):
        return False


def test_render_docx_accepts_single_record(monkeypatch, tmp_path: Path):
    fake_template = FakeTemplate()
    monkeypatch.setattr(render, "_open_template", lambda template: fake_template)

    result = render.render_docx(
        {"CN": "Acme"},
        tmp_path / "template.docx",
        tmp_path,
        filename=lambda m: f"{m['CN']}.docx",
    )

    assert result == tmp_path
    assert fake_template.doc.rendered == [{"CN": "Acme"}]
    assert fake_template.doc.saved == [tmp_path / "Acme.docx"]


def test_render_docx_rejects_output_outside_directory(monkeypatch, tmp_path: Path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    fake_template = FakeTemplate()
    monkeypatch.setattr(render, "_open_template", lambda template: fake_template)

    with pytest.raises(ValueError, match="ERR_OUTPUT_PATH") as exc_info:
        render.render_docx(
            {"CN": "Acme"},
            tmp_path / "template.docx",
            output_dir,
            filename=lambda record: "../escaped.docx",
        )

    assert not (tmp_path / "escaped.docx").exists()
    assert str(exc_info.value) == ascii_safe(exc_info.value)


def test_render_docx_rejects_duplicate_output_names(monkeypatch, tmp_path: Path):
    fake_template = FakeTemplate()
    create_output = Mock(return_value=tmp_path)
    monkeypatch.setattr(render, "_open_template", lambda template: fake_template)

    with pytest.raises(ValueError, match="ERR_OUTPUT_PATH"):
        render.render_docx(
            [{"CN": "Acme"}, {"CN": "ACME"}],
            tmp_path / "template.docx",
            create_output,
            filename=lambda record: f"{record['CN']}.docx",
        )

    create_output.assert_not_called()
    assert not hasattr(fake_template, "doc")


def test_open_template_wraps_permission_error(monkeypatch, tmp_path: Path):
    template = tmp_path / "locked.docx"
    template.write_text("template")

    def raise_permission_error(path):
        raise PermissionError("denied")

    monkeypatch.setattr(render, "DocxTemplate", raise_permission_error)

    with pytest.raises(RuntimeError) as exc_info:
        with render._open_template(template):
            pass

    message = str(exc_info.value)
    assert message.startswith("ERR_PERMISSION_DENIED:")
    assert f"template={ascii_safe(template.resolve())}" in message
    assert ascii_safe("无权限读取文件") in message
    assert message == ascii_safe(message)


class SavePermissionTemplate:
    def __init__(self, path):
        self.path = path

    def render(self, record):
        pass

    def save(self, path):
        raise PermissionError("output denied")


def test_render_docx_preserves_save_permission_error(monkeypatch, tmp_path: Path):
    template = tmp_path / "template.docx"
    template.write_text("template")

    monkeypatch.setattr(render, "DocxTemplate", SavePermissionTemplate)

    with pytest.raises(PermissionError) as exc_info:
        render.render_docx(
            {"CN": "Acme"},
            template,
            tmp_path,
            filename=lambda record: "Acme.docx",
        )

    assert str(exc_info.value) == "output denied"
