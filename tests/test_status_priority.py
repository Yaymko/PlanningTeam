"""Статус (справочник) и приоритет тикетов: перевод старой галочки «В работе», сортировка, перенос."""
import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import (  # noqa: E402
    DEFAULT_STATUSES,
    FIXED_COLS,
    merge_saved_hours,
    normalize_rows,
    sort_by_priority,
)
from storage import Store, rows_from_json  # noqa: E402

WEEKS = ["28.09–04.10 ч", "05.10–11.10 ч"]


def _row(ticket, priority="", status="", hours=0.0):
    return {"Тикет": ticket, "Название": f"Задача {ticket}", "Колонка": "10 Работа", "Доска": "1267",
            "Срок": None, "Статус": status, "Приоритет": priority, WEEKS[0]: hours, WEEKS[1]: 0.0}


def test_sort_by_priority_urgent_first_and_stable():
    df = pd.DataFrame([_row("A", "низкий"), _row("B"), _row("C", "срочно"), _row("D", "высокий"),
                       _row("E", "срочно"), _row("F", "средний")])
    assert sort_by_priority(df)["Тикет"].tolist() == ["C", "E", "D", "F", "A", "B"]


def test_legacy_in_work_column_maps_to_status():
    legacy = pd.DataFrame([{"Тикет": "A", "В работе": True, WEEKS[0]: 1.0},
                           {"Тикет": "B", "В работе": False, WEEKS[0]: 2.0}])
    df = normalize_rows(legacy, WEEKS)
    assert list(df.columns[:len(FIXED_COLS)]) == FIXED_COLS
    assert df["Статус"].tolist() == ["в работе", ""]
    assert df["Приоритет"].tolist() == ["", ""]


def test_rows_saved_before_status_column_load_with_status():
    old_json = json.dumps([{"Тикет": "A", "Название": "x", "Колонка": "", "Доска": "", "Срок": "2026-10-03",
                            "В работе": True, WEEKS[0]: 4.0, WEEKS[1]: 0.0}], ensure_ascii=False)
    df = rows_from_json(old_json, WEEKS)
    assert df.loc[0, "Статус"] == "в работе"
    assert df.loc[0, WEEKS[0]] == 4.0
    assert "В работе" not in df.columns
    assert rows_from_json("[]", WEEKS).empty


def test_fresh_board_upload_keeps_status_and_priority():
    old = pd.DataFrame([_row("A", "срочно", "пауза", 3.0), _row("B")])
    new = pd.DataFrame([_row("A", "", "в работе"), _row("B", "", "в работе")])
    merged = merge_saved_hours(new, old, WEEKS)
    assert merged["Статус"].tolist() == ["пауза", "в работе"]  # set in the planner wins, empty keeps the board's
    assert merged["Приоритет"].tolist() == ["срочно", ""]
    assert merged.loc[0, WEEKS[0]] == 3.0

    # even when no weeks overlap, status and priority are carried over
    other = merge_saved_hours(pd.DataFrame([_row("A")]).rename(columns={WEEKS[0]: "x"}), old, ["x"])
    assert (other.loc[0, "Статус"], other.loc[0, "Приоритет"]) == ("пауза", "срочно")


def test_save_person_sorts_by_priority(tmp_path):
    store = Store(tmp_path / "db.sqlite")
    store.save_plan("board.xlsx", {"Иванов": pd.DataFrame([_row("A", hours=1), _row("B", hours=1)])}, WEEKS)
    df, _, _ = store.load_person("Иванов")
    df.loc[1, "Приоритет"] = "срочно"
    store.save_person("Иванов", df, 40)
    df, _, _ = store.load_person("Иванов")
    assert df["Тикет"].tolist() == ["B", "A"]
    assert df.loc[0, "Приоритет"] == "срочно"


def test_status_reference(tmp_path):
    path = tmp_path / "db.sqlite"
    store = Store(path)
    assert store.statuses() == DEFAULT_STATUSES
    assert store.add_status("  на   ревью ")
    assert not store.add_status("На ревью")  # duplicate, case-insensitive
    assert not store.add_status("Пауза")
    assert not store.add_status("  ")
    assert Store(path).statuses() == [*DEFAULT_STATUSES, "на ревью"]
    store.remove_status("на ревью")
    store.remove_status("пауза")  # standard statuses stay
    assert store.statuses() == DEFAULT_STATUSES
