"""Alias entrypoint — Streamlit Cloud uses app.py."""
from pathlib import Path

_code = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
exec(compile(_code, str(Path(__file__).with_name("app.py")), "exec"), globals())
