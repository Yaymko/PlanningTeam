"""Smoke tests: the Streamlit app must start and render without exceptions,
both before a weekly plan is loaded and when employees open their hours."""
import sys
from pathlib import Path

import pandas as pd
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
APP_PATH = str(ROOT / "app.py")
sys.path.insert(0, str(ROOT))

from core import FIXED_COLS  # noqa: E402
from storage import Store  # noqa: E402

WEEK_LABELS = ["28.09–04.10 ч", "05.10–11.10 ч"]


def _plan_people():
    df = pd.DataFrame(
        [{"Тикет": "ABC-1", "Название": "Сделать штуку", "Колонка": "10 Работа", "Доска": "1267",
          "Срок": None, "В работе": True, WEEK_LABELS[0]: 4.0, WEEK_LABELS[1]: 0.0}],
        columns=FIXED_COLS + WEEK_LABELS,
    )
    return {"Иванов Иван": df, "Петрова Анна": df.copy()}


def _app(monkeypatch, db_path):
    monkeypatch.setenv("TEAM_LOAD_DB", str(db_path))
    return AppTest.from_file(APP_PATH, default_timeout=30)


def test_app_starts_without_plan(monkeypatch, tmp_path):
    at = _app(monkeypatch, tmp_path / "db.sqlite")
    at.run()

    assert not at.exception
    assert any("План недели ещё не загружен" in info.value for info in at.info)


def test_employee_saves_hours_and_planning_sees_them(monkeypatch, tmp_path):
    db = tmp_path / "db.sqlite"
    Store(db).save_plan("board.xlsx", _plan_people(), WEEK_LABELS)

    at = _app(monkeypatch, db)
    at.run()
    assert not at.exception
    assert any("Ещё не заполнили: Иванов Иван, Петрова Анна" in w.value for w in at.warning)

    at.selectbox[0].select("Иванов Иван").run()
    assert not at.exception
    at.number_input[0].set_value(32.0).run()
    next(b for b in at.button if b.label == "Сохранить").click().run()
    assert not at.exception

    _, norm, saved_at = Store(db).load_person("Иванов Иван")
    assert norm == 32.0
    assert saved_at
    assert any("Ещё не заполнили: Петрова Анна" in w.value for w in at.warning)
