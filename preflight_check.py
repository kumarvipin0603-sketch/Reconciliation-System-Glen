from __future__ import annotations

import ast
import pathlib
import py_compile
import sys

ROOT = pathlib.Path(__file__).resolve().parent
PY_FILES = [ROOT / "app.py", ROOT / "amazon_engine.py", ROOT / "flipkart_engine.py"]
REQUIRED_FILES = [
    *PY_FILES,
    ROOT / "requirements.txt",
    ROOT / "Pending_Task_Team_Working_Template.xlsx",
    ROOT / "Reconciliation_Pending_Task_Template.xlsx",
    ROOT / "Two_Dashboard_Format_Template.xlsx",
]

errors: list[str] = []

for path in REQUIRED_FILES:
    if not path.exists():
        errors.append(f"Missing required file: {path.name}")

for path in PY_FILES:
    if not path.exists():
        continue
    try:
        py_compile.compile(str(path), doraise=True)
    except Exception as exc:
        errors.append(f"Syntax/compile error in {path.name}: {exc}")

app_path = ROOT / "app.py"
if app_path.exists():
    source = app_path.read_text(encoding="utf-8")
    try:
        tree = ast.parse(source)
    except SyntaxError as exc:
        errors.append(f"AST parse error in app.py: {exc}")
        tree = None

    required_markers = {
        "Supabase DATABASE_URL support": "DATABASE_URL",
        "persistent source workbook table": "source_workbooks",
        "persistent task table": "pending_tasks",
        "persistent MIR table": "mir_details",
        "upload history": "upload_history",
        "source restore": "def saved_source_path(portal):",
        "team history preservation": "def restore_operational_history(history):",
    }
    for label, marker in required_markers.items():
        if marker not in source:
            errors.append(f"Missing {label} marker: {marker}")

    # Prevent the exact production error previously seen in MIR team upload.
    if "log_activity(\n                \"TEAM UPDATE\",\n                branch=branch" in source:
        errors.append("Invalid log_activity keyword found: branch=branch; use branch_code=branch")

    # Guardrail: source refresh must not auto-write the historic team completion phrase.
    if "Auto-closed after reconciliation status changed" in source:
        errors.append("Source refresh still contains automatic team-task closure logic")

    if tree is not None:
        defs = {
            n.name: n
            for n in ast.walk(tree)
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
        }
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Name):
                continue
            fn = defs.get(node.func.id)
            if not fn or fn.args.kwarg is not None:
                continue
            allowed = {a.arg for a in fn.args.args + fn.args.kwonlyargs}
            for kw in node.keywords:
                if kw.arg and kw.arg not in allowed:
                    errors.append(
                        f"Unexpected keyword '{kw.arg}' passed to {node.func.id} at app.py:{node.lineno}"
                    )

if errors:
    print("\nPRE-FLIGHT FAILED\n")
    for item in errors:
        print(f" - {item}")
    print("\nNothing should be deployed until the above issue(s) are fixed.")
    sys.exit(1)

print("PRE-FLIGHT PASSED")
print(" - Python files compile")
print(" - Required project files exist")
print(" - Persistence guardrails are present")
print(" - No known branch/branch_code logging mismatch found")
