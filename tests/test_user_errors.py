import ast
import inspect
import re
from pathlib import Path


ERROR_CODE_PATTERN = re.compile(r"\bERR_[A-Z0-9_]+\b")
PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _production_paths() -> list[Path]:
    return [
        *sorted((PROJECT_ROOT / "src" / "sheetflow").rglob("*.py")),
        *sorted((PROJECT_ROOT / "py_script").rglob("*.py")),
    ]


def _registered_error_codes() -> set[str]:
    from sheetflow import user_errors

    return {
        name
        for name, value in vars(user_errors).items()
        if name.startswith("ERR_") and value == name
    }


def _format_user_error_code_names(source: str) -> set[str]:
    tree = ast.parse(source)
    used = set()
    for node in ast.walk(tree):
        if not _is_format_user_error_call(node):
            continue
        code = node.args[0] if node.args else None
        if isinstance(code, ast.Name):
            used.add(code.id)
    return used


def _is_format_user_error_call(node: ast.AST) -> bool:
    if not isinstance(node, ast.Call):
        return False
    return (
        isinstance(node.func, ast.Name) and node.func.id == "format_user_error"
    ) or (
        isinstance(node.func, ast.Attribute)
        and node.func.attr == "format_user_error"
    )


def test_common_reexports_user_error_helpers():
    from sheetflow import common, user_errors

    assert common.ascii_safe is user_errors.ascii_safe
    assert common.format_user_error is user_errors.format_user_error


def test_user_errors_has_no_project_dependencies():
    from sheetflow import user_errors

    tree = ast.parse(inspect.getsource(user_errors))
    relative_imports = [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.ImportFrom) and node.level
    ]

    assert relative_imports == []


def test_smart_path_manager_does_not_copy_ascii_safe():
    from sheetflow import smart_path_manager

    assert not hasattr(smart_path_manager, "_ascii_safe")


def test_error_code_registry_matches_production_usage():
    used = set()
    for path in _production_paths():
        if path.name == "user_errors.py":
            continue
        used.update(
            _format_user_error_code_names(path.read_text(encoding="utf-8"))
        )

    assert _registered_error_codes() == used


def test_production_usage_ignores_error_codes_outside_formatter_calls():
    source = '''
"""Mentions ERR_STALE_DOCSTRING without using it."""
ERR_UNUSED_IMPORT = "ERR_UNUSED_IMPORT"
format_user_error(ERR_REAL, "english", "chinese")
'''

    assert _format_user_error_code_names(source) == {"ERR_REAL"}


def test_readme_documents_every_registered_error_code():
    documented = set(
        ERROR_CODE_PATTERN.findall(
            (PROJECT_ROOT / "README.md").read_text(encoding="utf-8")
        )
    )

    assert documented == _registered_error_codes()


def test_format_user_error_calls_use_registered_constants():
    registered = _registered_error_codes()
    invalid_calls = []

    for path in _production_paths():
        if path.name == "user_errors.py":
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not _is_format_user_error_call(node):
                continue

            code = node.args[0] if node.args else None
            if not isinstance(code, ast.Name) or code.id not in registered:
                invalid_calls.append(f"{path.relative_to(PROJECT_ROOT)}:{node.lineno}")

    assert invalid_calls == []
