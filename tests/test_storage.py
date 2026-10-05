"""Shared storage: hours saved by one employee survive a restart and a fresh board upload."""
import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import FIXED_COLS  # noqa: E402
from storage import Store  # noqa: E402

WEEKS = ["28.09–04.10 ч", "05.10–11.10 ч"]
NEXT_WEEKS = ["05.10–11.10 ч", "12.10–18.10 ч"]


def _df(weeks, rows):
    return pd.DataFrame(rows, columns=FIXED_COLS + weeks)


def _ticket(key, weeks, hours, due=None):
    return {"Тикет": key, "Название": f"Задача {key}", "Колонка": "10 Работа", "Доска": "1267",
            "Срок": due, "В работе": True, **dict(zip(weeks, hours))}


def test_no_plan_initially(tmp_path):
    store = Store(tmp_path / "db.sqlite")
    assert store.get_plan() is None
    assert store.people() == []
    assert store.load_all() == {}


def test_saved_hours_survive_restart(tmp_path):
    path = tmp_path / "db.sqlite"
    store = Store(path)
    store.save_plan("board.xlsx", {"Иванов": _df(WEEKS, [_ticket("A-1", WEEKS, [0, 0], date(2026, 10, 3))])}, WEEKS)

    df, norm, saved_at = store.load_person("Иванов")
    assert norm == 40.0 and saved_at is None
    df.loc[0, WEEKS[0]] = 6.5
    extra = {"Тикет": None, "Название": "Созвоны", "Колонка": None, "Доска": None, "Срок": None,
             "В работе": None, WEEKS[0]: 2, WEEKS[1]: None}  # a row added in data_editor
    store.save_person("Иванов", pd.concat([df, pd.DataFrame([extra])], ignore_index=True), 36)

    df, norm, saved_at = Store(path).load_person("Иванов")  # "restart"
    assert norm == 36.0 and saved_at
    assert df.loc[0, WEEKS[0]] == 6.5
    assert df.loc[0, "Срок"] == date(2026, 10, 3)
    assert df.loc[1, "Название"] == "Созвоны"
    assert df.loc[1, "Колонка"] == "Доп. работа"
    assert df.loc[1, WEEKS[1]] == 0.0


def test_reupload_keeps_hours_for_matching_tickets_and_weeks(tmp_path):
    store = Store(tmp_path / "db.sqlite")
    store.save_plan("w1.xlsx", {"Иванов": _df(WEEKS, [_ticket("A-1", WEEKS, [0, 0])])}, WEEKS)
    df, _, _ = store.load_person("Иванов")
    df.loc[0, WEEKS] = [5.0, 7.0]
    extra = _df(WEEKS, [{"Тикет": "", "Название": "Созвоны", "Колонка": "Доп. работа", "Доска": "",
                         "Срок": None, "В работе": False, WEEKS[0]: 1.0, WEEKS[1]: 3.0}])
    store.save_person("Иванов", pd.concat([df, extra], ignore_index=True), 30)

    # Same weeks re-uploaded: everything kept, including the "filled" mark.
    store.save_plan("w1b.xlsx", {"Иванов": _df(WEEKS, [_ticket("A-1", WEEKS, [0, 0])])}, WEEKS)
    df, norm, saved_at = store.load_person("Иванов")
    assert list(df[WEEKS[0]]) == [5.0, 1.0] and norm == 30.0 and saved_at

    # Next week's export: overlapping week carried over, new week empty, "filled" mark reset.
    new = _df(NEXT_WEEKS, [_ticket("A-1", NEXT_WEEKS, [0, 0]), _ticket("B-2", NEXT_WEEKS, [0, 0])])
    store.save_plan("w2.xlsx", {"Иванов": new, "Петрова": new.copy()}, NEXT_WEEKS)
    df, norm, saved_at = store.load_person("Иванов")
    assert list(df["Тикет"]) == ["A-1", "B-2", ""]
    assert list(df[NEXT_WEEKS[0]]) == [7.0, 0.0, 3.0]
    assert list(df[NEXT_WEEKS[1]]) == [0.0, 0.0, 0.0]
    assert norm == 30.0 and saved_at is None
    assert store.people() == ["Иванов", "Петрова"]


def test_rows_marked_for_deletion_are_removed_on_save(tmp_path):
    from core import DELETE_COL, drop_marked_rows

    store = Store(tmp_path / "db.sqlite")
    rows = [_ticket("A-1", WEEKS, [1, 0]), _ticket("B-2", WEEKS, [2, 0]), _ticket("C-3", WEEKS, [3, 0])]
    store.save_plan("w1.xlsx", {"Иванов": _df(WEEKS, rows)}, WEEKS)

    df, _, _ = store.load_person("Иванов")
    df.insert(0, DELETE_COL, [False, True, None])  # None: a row added in data_editor, box untouched
    store.save_person("Иванов", drop_marked_rows(df), 40)

    df, _, _ = store.load_person("Иванов")
    assert list(df["Тикет"]) == ["A-1", "C-3"]
    assert DELETE_COL not in df.columns


def test_actual_hours_saved_and_carried_over_on_reupload(tmp_path):
    from core import actual_label

    store = Store(tmp_path / "db.sqlite")
    store.save_plan("w1.xlsx", {"Иванов": _df(WEEKS, [_ticket("A-1", WEEKS, [5, 7])])}, WEEKS)
    df, _, _ = store.load_person("Иванов")
    assert df.loc[0, actual_label(WEEKS[0])] == 0.0  # board export has no actuals
    df.loc[0, actual_label(WEEKS[0])] = 6.0
    store.save_person("Иванов", df, 40)

    df, _, _ = store.load_person("Иванов")
    assert df.loc[0, actual_label(WEEKS[0])] == 6.0

    new = _df(NEXT_WEEKS, [_ticket("A-1", NEXT_WEEKS, [0, 0])])
    store.save_plan("w2.xlsx", {"Иванов": new}, NEXT_WEEKS)
    df, _, _ = store.load_person("Иванов")
    assert df.loc[0, NEXT_WEEKS[0]] == 7.0
    assert df.loc[0, actual_label(NEXT_WEEKS[0])] == 0.0
    assert actual_label(WEEKS[0]) not in df.columns


def test_database_saved_before_actual_hours_still_loads(tmp_path):
    """A team_load.db written by the previous version (rows JSON without actual hours) keeps working."""
    import json
    import sqlite3

    from core import actual_label

    path = tmp_path / "db.sqlite"
    Store(path)
    old_rows = [{"Тикет": "A-1", "Название": "Задача", "Колонка": "10 Работа", "Доска": "1267",
                 "Срок": "2026-10-03", "В работе": True, WEEKS[0]: 4.0, WEEKS[1]: 2.0}]
    with sqlite3.connect(path) as conn:
        conn.execute("INSERT INTO plan (id, file_name, uploaded_at, week_labels) VALUES (1, 'old.xlsx', "
                     "'2026-10-01T10:00:00+03:00', ?)", (json.dumps(WEEKS, ensure_ascii=False),))
        conn.execute("INSERT INTO people (name, position, rows, norm, updated_at) VALUES (?, 0, ?, 32, NULL)",
                     ("Иванов", json.dumps(old_rows, ensure_ascii=False)))

    store = Store(path)
    df, norm, _ = store.load_person("Иванов")
    assert df.loc[0, WEEKS[0]] == 4.0 and norm == 32.0
    assert df.loc[0, actual_label(WEEKS[1])] == 0.0
    df.loc[0, actual_label(WEEKS[1])] = 3.5
    store.save_person("Иванов", df, norm)
    assert store.load_person("Иванов")[0].loc[0, actual_label(WEEKS[1])] == 3.5
