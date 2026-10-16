"""Поиск и фильтры таблиц: фильтрация строк и перенос правок из отфильтрованного редактора."""
import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from table_filter import EMPTY, apply_filter, merge_view_edits, search_mask  # noqa: E402


def _df():
    return pd.DataFrame({
        "Тикет": ["ABC-1", "ABC-2", "XYZ-3"],
        "Название": ["Сделать отчёт", "Починить импорт", None],
        "Срок": pd.to_datetime(["2026-10-05", None, "2026-10-20"]),
        "В работе": [True, False, True],
        "Часы": [4.0, 0.5, 10.0],
    })


def test_search_is_case_insensitive_across_columns_and_words():
    df = _df()
    assert search_mask(df, "abc").tolist() == [True, True, False]
    assert search_mask(df, "ИМПОРТ").tolist() == [False, True, False]
    assert search_mask(df, "abc отчёт").tolist() == [True, False, False]  # all words in one row
    assert search_mask(df, "20.10.2026").tolist() == [False, False, True]  # dates as shown
    assert search_mask(df, "0.5").tolist() == [False, True, False]
    assert search_mask(df, "  ").all()


def test_column_filters():
    df = _df()
    assert apply_filter(df, {"filters": {"Тикет": ("values", ["XYZ-3", "ABC-1"])}})["Тикет"].tolist() == \
        ["ABC-1", "XYZ-3"]
    assert apply_filter(df, {"filters": {"Название": ("values", [EMPTY])}})["Тикет"].tolist() == ["XYZ-3"]
    assert apply_filter(df, {"filters": {"В работе": ("values", ["нет"])}})["Тикет"].tolist() == ["ABC-2"]
    assert apply_filter(df, {"filters": {"Часы": ("range", (1, None))}})["Тикет"].tolist() == ["ABC-1", "XYZ-3"]
    assert apply_filter(df, {"filters": {"Часы": ("range", (None, 5))}})["Тикет"].tolist() == ["ABC-1", "ABC-2"]
    period = (date(2026, 10, 1), date(2026, 10, 5))
    assert apply_filter(df, {"filters": {"Срок": ("dates", period)}})["Тикет"].tolist() == ["ABC-1"]
    assert apply_filter(df, {"filters": {"Название": ("contains", "ПОЧИН")}})["Тикет"].tolist() == ["ABC-2"]
    # empty filter values show everything; search and filters combine
    assert len(apply_filter(df, {"filters": {"Тикет": ("values", [])}})) == 3
    spec = {"query": "abc", "filters": {"Часы": ("range", (1, None))}}
    assert apply_filter(df, spec)["Тикет"].tolist() == ["ABC-1"]


def test_merge_view_edits_keeps_hidden_rows_and_order():
    full = _df().assign(_id=range(3))
    view = full[full["Тикет"] != "ABC-1"]  # rows 1 and 2 shown
    edited = view.copy()
    edited.loc[edited["_id"] == 2, "Часы"] = 7.0
    edited = edited[edited["_id"] != 1]  # row deleted in the editor
    added = pd.DataFrame([{"Тикет": "", "Название": "Доп. работа", "Часы": 3.0, "_id": None}])
    edited = pd.concat([edited, added], ignore_index=True)

    merged = merge_view_edits(full, view, edited, "_id")
    assert merged["Тикет"].tolist() == ["ABC-1", "XYZ-3", ""]
    assert merged["Часы"].tolist() == [4.0, 7.0, 3.0]
    assert merged["_id"].tolist() == [0, 2, 3]
