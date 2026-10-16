"""
Поиск и фильтры для таблиц приложения.

Над таблицей — строка поиска (ищет текст во всех колонках) и выбор колонок для фильтра;
для каждой выбранной колонки появляется свой фильтр: список значений, диапазон чисел или дат.
Виджеты рисует filter_controls, сама фильтрация — чистая функция apply_filter.
"""
from datetime import date

import pandas as pd
import streamlit as st

EMPTY = "(пусто)"
MAX_CHOICES = 60  # больше уникальных значений — фильтр «содержит» вместо списка


def _kind(s: pd.Series) -> str:
    if pd.api.types.is_bool_dtype(s):
        return "bool"
    if pd.api.types.is_datetime64_any_dtype(s):
        return "date"
    if pd.api.types.is_numeric_dtype(s):
        return "number"
    non_null = s.dropna()
    if not non_null.empty and non_null.map(lambda v: isinstance(v, bool)).all():
        return "bool"
    return "text"


def _cell_text(v) -> str:
    if v is None or (not isinstance(v, (list, tuple)) and pd.isna(v)):
        return ""
    if isinstance(v, bool):
        return "да" if v else "нет"
    if isinstance(v, (pd.Timestamp, date)):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float):
        return f"{v:g}"
    return str(v)


def _choice(v) -> str:
    text = _cell_text(v)
    return text if text else EMPTY


def search_mask(df: pd.DataFrame, query: str, columns=None) -> pd.Series:
    """Строки, где каждое слово запроса встречается хотя бы в одной колонке (без учёта регистра)."""
    words = query.lower().split()
    mask = pd.Series(True, index=df.index)
    if not words or df.empty:
        return mask
    cols = list(columns) if columns is not None else list(df.columns)
    texts = df[cols].apply(lambda col: col.map(_cell_text).str.lower())
    for word in words:
        mask &= texts.apply(lambda col: col.str.contains(word, regex=False)).any(axis=1)
    return mask


def apply_filter(df: pd.DataFrame, spec: dict) -> pd.DataFrame:
    """spec: {"query": str, "columns": [...] , "filters": {колонка: (вид, значение)}}."""
    mask = search_mask(df, spec.get("query", ""), spec.get("columns"))
    for col, (kind, value) in spec.get("filters", {}).items():
        if col not in df.columns or value in (None, [], ()):
            continue
        s = df[col]
        if kind == "values":
            mask &= s.map(_choice).isin(value)
        elif kind == "contains":
            mask &= s.map(_cell_text).str.lower().str.contains(value.lower(), regex=False)
        elif kind == "range":
            lo, hi = value
            num = pd.to_numeric(s, errors="coerce")
            if lo is not None:
                mask &= num >= lo
            if hi is not None:
                mask &= num <= hi
        elif kind == "dates":
            lo, hi = value
            d = pd.to_datetime(s, errors="coerce")
            mask &= (d >= pd.Timestamp(lo)) & (d < pd.Timestamp(hi) + pd.Timedelta(days=1))
    return df[mask]


def filter_controls(df: pd.DataFrame, key: str, columns=None, labels=None) -> dict:
    """Рисует поиск и фильтры над таблицей и возвращает их состояние для apply_filter."""
    labels = labels or {}
    cols = [c for c in (columns if columns is not None else df.columns) if c in df.columns]
    left, right = st.columns([2, 3])
    query = left.text_input("Поиск по таблице", key=f"{key}_search", placeholder="Текст в любой колонке",
                            help="Ищет в каждой колонке без учёта регистра; несколько слов — все должны найтись "
                                 "в строке.")
    chosen = right.multiselect("Фильтр по колонкам", cols, key=f"{key}_cols",
                               format_func=lambda c: labels.get(c, c), placeholder="Выберите колонки")
    filters = {}
    if chosen:
        boxes = st.columns(min(len(chosen), 3))
        for i, col in enumerate(chosen):
            name = labels.get(col, col)
            s = df[col]
            kind = _kind(s)
            with boxes[i % len(boxes)]:
                wkey = f"{key}_f_{col}"
                if kind == "number":
                    a, b = st.columns(2)
                    lo = a.number_input(f"{name}: от", value=None, key=f"{wkey}_lo", format="%g")
                    hi = b.number_input(f"{name}: до", value=None, key=f"{wkey}_hi", format="%g")
                    filters[col] = ("range", (lo, hi))
                elif kind == "date":
                    picked = st.date_input(f"{name}: период", value=(), key=wkey, format="DD.MM.YYYY")
                    if isinstance(picked, (list, tuple)) and len(picked) == 2:
                        filters[col] = ("dates", tuple(picked))
                    elif isinstance(picked, (list, tuple)) and len(picked) == 1:
                        filters[col] = ("dates", (picked[0], picked[0]))
                else:
                    options = sorted(set(s.map(_choice)), key=lambda v: (v == EMPTY, v.lower()))
                    if kind == "text" and len(options) > MAX_CHOICES:
                        value = st.text_input(f"{name}: содержит", key=wkey)
                        filters[col] = ("contains", value.strip()) if value.strip() else ("values", [])
                    else:
                        filters[col] = ("values", st.multiselect(name, options, key=wkey,
                                                                 placeholder="Все значения"))
    return {"query": query, "columns": cols, "filters": filters}


def filter_table(df: pd.DataFrame, key: str, columns=None, labels=None) -> pd.DataFrame:
    """Поиск и фильтры над таблицей; возвращает отфильтрованные строки и пишет, сколько показано."""
    shown = apply_filter(df, filter_controls(df, key, columns, labels))
    if len(shown) != len(df):
        st.caption(f"Показано строк: {len(shown)} из {len(df)}")
    return shown


def merge_view_edits(full: pd.DataFrame, view: pd.DataFrame, edited: pd.DataFrame, id_col: str) -> pd.DataFrame:
    """
    Возвращает полную таблицу с правками, сделанными в отфильтрованном редакторе.
    Строки вне фильтра остаются как были, показанные заменяются отредактированными (удалённые
    в редакторе пропадают), добавленные дописываются в конец с новыми id.
    """
    shown_ids = set(view[id_col])
    ids = pd.to_numeric(edited[id_col], errors="coerce")
    existing = edited[ids.isin(shown_ids)]
    added = edited[ids.isna()].copy()
    hidden = full[~full[id_col].isin(shown_ids)]
    order = {rid: pos for pos, rid in enumerate(full[id_col])}
    merged = pd.concat([hidden, existing])
    merged = merged.iloc[merged[id_col].map(order).argsort(kind="stable")]
    next_id = int(full[id_col].max()) + 1 if len(full) else 0
    added[id_col] = range(next_id, next_id + len(added))
    merged = pd.concat([merged, added], ignore_index=True)
    merged[id_col] = merged[id_col].astype(int)
    return merged
