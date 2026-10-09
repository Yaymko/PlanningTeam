"""
Общее хранилище плана недели и часов сотрудников (SQLite-файл).

Без зависимостей от Streamlit. Путь к файлу базы задаётся переменной окружения
TEAM_LOAD_DB, по умолчанию — data/team_load.db рядом с приложением.
"""
import json
import os
import sqlite3
from contextlib import closing
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from core import DEFAULT_NORM, FIXED_COLS, actual_label, hour_cols, merge_saved_hours, normalize_rows, week_key

DEFAULT_DB_PATH = Path(__file__).resolve().parent / "data" / "team_load.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS plan (
    id          INTEGER PRIMARY KEY CHECK (id = 1),
    file_name   TEXT NOT NULL,
    uploaded_at TEXT NOT NULL,
    week_labels TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS people (
    name       TEXT PRIMARY KEY,
    position   INTEGER NOT NULL,
    rows       TEXT NOT NULL,
    norm       REAL NOT NULL,
    updated_at TEXT
);
-- Прошедшие недели, которые ушли из плана при загрузке новой выгрузки: часы каждого сотрудника
-- (план и факт) за одну неделю, норма и когда сотрудник сохранял.
CREATE TABLE IF NOT EXISTS week_history (
    week_key    TEXT NOT NULL,
    week_label  TEXT NOT NULL,
    name        TEXT NOT NULL,
    position    INTEGER NOT NULL,
    rows        TEXT NOT NULL,
    norm        REAL NOT NULL,
    updated_at  TEXT,
    file_name   TEXT NOT NULL,
    archived_at TEXT NOT NULL,
    PRIMARY KEY (week_key, name)
);
"""


def db_path_from_env() -> Path:
    return Path(os.environ.get("TEAM_LOAD_DB") or DEFAULT_DB_PATH)


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _date_to_str(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    parsed = pd.to_datetime(v, errors="coerce", dayfirst=True)
    return None if pd.isna(parsed) else parsed.date().isoformat()


def rows_to_json(df: pd.DataFrame, week_labels: list) -> str:
    df = normalize_rows(df, week_labels)
    records = []
    for _, row in df.iterrows():
        rec = {col: row[col] for col in FIXED_COLS}
        rec["Срок"] = _date_to_str(rec["Срок"])
        rec["В работе"] = bool(rec["В работе"])
        for col in hour_cols(week_labels):
            rec[col] = float(row[col])
        records.append(rec)
    return json.dumps(records, ensure_ascii=False)


def rows_from_json(text: str, week_labels: list) -> pd.DataFrame:
    # Строки, сохранённые до появления фактических часов, их не содержат — normalize_rows ставит 0.
    df = pd.DataFrame(json.loads(text), columns=FIXED_COLS + hour_cols(week_labels))
    df["Срок"] = [date.fromisoformat(v) if isinstance(v, str) and v else None for v in df["Срок"]]
    return normalize_rows(df, week_labels)


class Store:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with closing(self._connect()) as conn, conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.executescript(SCHEMA)

    def _connect(self):
        return sqlite3.connect(self.path, timeout=30)

    # --- план недели -----------------------------------------------------------------

    def get_plan(self):
        """Текущий план: {'file_name', 'uploaded_at', 'week_labels'} или None, если не загружен."""
        with closing(self._connect()) as conn:
            row = conn.execute("SELECT file_name, uploaded_at, week_labels FROM plan WHERE id = 1").fetchone()
        if row is None:
            return None
        return {"file_name": row[0], "uploaded_at": row[1], "week_labels": json.loads(row[2])}

    def save_plan(self, file_name: str, people_data: dict, week_labels: list, today: date = None) -> None:
        """Делает свежую выгрузку с доски планом недели.

        Уже введённые часы по совпадающим тикетам и неделям переносятся, нормы сохраняются.
        Если набор недель поменялся (новая неделя), отметки «заполнено» сбрасываются.
        Недели, которых нет в новой выгрузке, уходят в историю; если неделя из истории снова
        появилась в выгрузке, её часы возвращаются в план.
        """
        today = today or date.today()
        old_plan = self.get_plan()
        old_weeks = old_plan["week_labels"] if old_plan else []
        same_weeks = list(old_weeks) == list(week_labels)
        dropped = [wl for wl in old_weeks if wl not in week_labels]
        returned = {week_key(wl, today): wl for wl in week_labels if wl not in old_weeks}
        with closing(self._connect()) as conn, conn:
            old = {
                name: (rows, norm, updated_at, position)
                for name, rows, norm, updated_at, position in conn.execute(
                    "SELECT name, rows, norm, updated_at, position FROM people")
            }
            for name, (old_rows, norm, updated_at, position) in old.items():
                old_df = rows_from_json(old_rows, old_weeks)
                for wl in dropped:
                    week_df = old_df[FIXED_COLS + hour_cols([wl])]
                    week_df = week_df[(week_df[wl] != 0) | (week_df[actual_label(wl)] != 0)]
                    conn.execute(
                        "INSERT OR REPLACE INTO week_history (week_key, week_label, name, position, rows, norm, "
                        "updated_at, file_name, archived_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (week_key(wl, today), wl, name, position, rows_to_json(week_df, [wl]), norm,
                         updated_at, old_plan["file_name"], _now()))
            history = {}
            if returned:
                marks = ",".join("?" * len(returned))
                for key, name, rows, label in conn.execute(
                        f"SELECT week_key, name, rows, week_label FROM week_history WHERE week_key IN ({marks})",
                        list(returned)):
                    history.setdefault(name, []).append(
                        rows_from_json(rows, [label]).rename(columns={
                            label: returned[key], actual_label(label): actual_label(returned[key])}))
                conn.execute(f"DELETE FROM week_history WHERE week_key IN ({marks})", list(returned))
            conn.execute("DELETE FROM people")
            for position, (name, df) in enumerate(people_data.items()):
                norm, updated_at = DEFAULT_NORM, None
                if name in old:
                    old_rows, norm, old_updated, _ = old[name]
                    df = merge_saved_hours(df, rows_from_json(old_rows, old_weeks), week_labels)
                    updated_at = old_updated if same_weeks else None
                for hist_df in history.get(name, []):
                    df = merge_saved_hours(df, hist_df, week_labels)
                conn.execute(
                    "INSERT INTO people (name, position, rows, norm, updated_at) VALUES (?, ?, ?, ?, ?)",
                    (name, position, rows_to_json(df, week_labels), norm, updated_at))
            conn.execute(
                "INSERT OR REPLACE INTO plan (id, file_name, uploaded_at, week_labels) VALUES (1, ?, ?, ?)",
                (file_name, _now(), json.dumps(list(week_labels), ensure_ascii=False)))

    # --- часы сотрудников ------------------------------------------------------------

    def people(self) -> list:
        with closing(self._connect()) as conn:
            return [r[0] for r in conn.execute("SELECT name FROM people ORDER BY position")]

    def load_person(self, name: str):
        """(таблица тикетов, норма, когда сохранено) для сотрудника."""
        plan = self.get_plan()
        with closing(self._connect()) as conn:
            row = conn.execute("SELECT rows, norm, updated_at FROM people WHERE name = ?", (name,)).fetchone()
        if row is None or plan is None:
            raise KeyError(name)
        return rows_from_json(row[0], plan["week_labels"]), row[1], row[2]

    def load_all(self) -> dict:
        """{имя: (таблица тикетов, норма, когда сохранено)} в порядке листов выгрузки."""
        plan = self.get_plan()
        if plan is None:
            return {}
        with closing(self._connect()) as conn:
            rows = conn.execute("SELECT name, rows, norm, updated_at FROM people ORDER BY position").fetchall()
        return {name: (rows_from_json(r, plan["week_labels"]), norm, upd) for name, r, norm, upd in rows}

    def save_person(self, name: str, df: pd.DataFrame, norm: float) -> None:
        plan = self.get_plan()
        if plan is None:
            raise RuntimeError("План недели ещё не загружен")
        with closing(self._connect()) as conn, conn:
            cur = conn.execute(
                "UPDATE people SET rows = ?, norm = ?, updated_at = ? WHERE name = ?",
                (rows_to_json(df, plan["week_labels"]), float(norm), _now(), name))
            if cur.rowcount == 0:
                raise KeyError(name)

    # --- история прошедших недель ----------------------------------------------------

    def history_weeks(self) -> list:
        """Недели в истории, новые первыми: [{'key', 'label', 'file_name', 'archived_at'}]."""
        with closing(self._connect()) as conn:
            rows = conn.execute(
                "SELECT week_key, MAX(week_label), MAX(file_name), MAX(archived_at) FROM week_history "
                "GROUP BY week_key ORDER BY week_key DESC").fetchall()
        return [{"key": k, "label": lbl, "file_name": f, "archived_at": a} for k, lbl, f, a in rows]

    def load_history_week(self, key: str) -> tuple:
        """(заголовок недели, {имя: (таблица тикетов, норма, когда сохранено)}) для недели из истории."""
        with closing(self._connect()) as conn:
            rows = conn.execute(
                "SELECT week_label, name, rows, norm, updated_at FROM week_history WHERE week_key = ? "
                "ORDER BY position, name", (key,)).fetchall()
        if not rows:
            raise KeyError(key)
        label = rows[0][0]
        return label, {name: (rows_from_json(r, [label]), norm, upd) for _, name, r, norm, upd in rows}
