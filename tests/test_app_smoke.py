"""Smoke test: the Streamlit app must start and render without exceptions
when no file has been uploaded yet."""
import sys
from pathlib import Path

from streamlit.testing.v1 import AppTest

APP_PATH = str(Path(__file__).resolve().parents[1] / "app.py")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def test_app_starts_without_upload():
    at = AppTest.from_file(APP_PATH)
    at.run(timeout=30)

    assert not at.exception
    assert any("Загрузите еженедельный экспорт" in info.value for info in at.info)
