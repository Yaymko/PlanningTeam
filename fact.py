"""
Фактические часы из выгрузки учёта рабочего времени («Сводная таблица по выполненным работам»)
и сравнение их с планом из планировщика. Без зависимостей от Streamlit.
"""
import io
import re
from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd

from core import actual_label, week_key

NAME_COL = "Фамилия И.О."
DATE_COL = "Дата"
HOURS_COL = "Часы"
NOTE_COL = "Примечание"
# Колонки выгрузки, по которым можно разложить фактические часы; первая — группировка по умолчанию.
GROUP_COLS = ["Этап внутреннего плана", "Внутренний план работ", "Проект", "Вид затрат",
              "Этап проекта", "Тема расчетов", "Направление деятельности"]
FACT_COLS = [DATE_COL, "Сотрудник", *GROUP_COLS, HOURS_COL, NOTE_COL]
NO_VALUE = "(не указано)"


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    parsed = pd.to_datetime(v, errors="coerce", dayfirst=True)
    return None if pd.isna(parsed) else parsed.date()


def parse_fact_workbook(file_bytes: bytes) -> pd.DataFrame:
    """Разбирает выгрузку учёта времени: строка = сотрудник, день, работа, часы.

    Лист и строка заголовка ищутся по колонкам «Фамилия И.О.» и «Часы». Строки без сотрудника,
    без даты или с 0 ч пропускаются. Возвращает DataFrame с колонками FACT_COLS.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    records = []
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        idx = None
        for _ in range(10):
            header = next(rows, None)
            if header is None:
                break
            names = [str(h).strip() if h is not None else "" for h in header]
            if NAME_COL in names and HOURS_COL in names and DATE_COL in names:
                idx = {h: i for i, h in enumerate(names) if h}
                break
        if idx is None:
            continue
        for row in rows:
            def get(col):
                i = idx.get(col)
                return row[i] if i is not None and i < len(row) else None
            name, day, hours = get(NAME_COL), _to_date(get(DATE_COL)), get(HOURS_COL)
            if not name or day is None:
                continue
            hours = pd.to_numeric(hours, errors="coerce")
            if pd.isna(hours) or hours == 0:
                continue
            rec = {DATE_COL: day, "Сотрудник": " ".join(str(name).split()), HOURS_COL: float(hours),
                   NOTE_COL: str(get(NOTE_COL) or "").strip()}
            for col in GROUP_COLS:
                rec[col] = str(get(col) or "").strip() or NO_VALUE
            records.append(rec)
    wb.close()
    return pd.DataFrame(records, columns=FACT_COLS)


def monday(d: date) -> date:
    return d - timedelta(days=d.weekday())


def split_by_week(fact_df: pd.DataFrame) -> dict:
    """{ключ недели (дата понедельника, YYYY-MM-DD): строки факта за эту неделю}."""
    if fact_df.empty:
        return {}
    keys = fact_df[DATE_COL].map(lambda d: monday(d).isoformat())
    return {k: part.reset_index(drop=True) for k, part in fact_df.groupby(keys, sort=True)}


def week_title(key: str) -> str:
    """«2026-10-12» -> «12.10–18.10»."""
    start = date.fromisoformat(key)
    return f"{start:%d.%m}–{start + timedelta(days=6):%d.%m}"


# --- сопоставление сотрудников ---------------------------------------------------------

def _name_parts(name: str) -> list:
    return [p for p in re.split(r"[\s.]+", str(name).lower().replace("ё", "е")) if p]


def _same_person(a: str, b: str) -> bool:
    """Фамилия совпадает, имя и отчество — если указаны у обоих — совпадают или одно из них инициал.
    «Иванов Иван» = «Иванов Иван Петрович», «Кузнецов А.Б.» = «Кузнецов Андрей Борисович»."""
    pa, pb = _name_parts(a), _name_parts(b)
    if not pa or not pb or pa[0] != pb[0]:
        return False
    for x, y in zip(pa[1:3], pb[1:3]):
        if x != y and not ((len(x) == 1 or len(y) == 1) and x[0] == y[0]):
            return False
    return True


def match_people(fact_names, plan_names) -> dict:
    """{имя в выгрузке факта: имя в планировщике}. Сначала точное совпадение, затем по фамилии
    и имени/инициалам, если подходит ровно один сотрудник планировщика."""
    plan_names = list(plan_names)
    result = {}
    for name in fact_names:
        if name in plan_names:
            result[name] = name
            continue
        candidates = [p for p in plan_names if _same_person(name, p)]
        if len(candidates) == 1:
            result[name] = candidates[0]
    return result


# --- сравнение план / факт ---------------------------------------------------------------

def plan_week_label(week_labels, key: str, today: date):
    """Заголовок недели плана с этим ключом (датой понедельника) или None."""
    return next((wl for wl in week_labels if week_key(wl, today) == key), None)


def compare_plan_fact(plan_people: dict, label, fact_df: pd.DataFrame) -> tuple:
    """Сводка по сотрудникам за неделю.

    plan_people — {имя: (таблица тикетов, норма, когда сохранено)} из планировщика; label — колонка
    недели в этих таблицах (None, если плана на неделю нет). Возвращает (сводка, факт с колонкой
    «Сотрудник» в именах планировщика). Сотрудники, которые есть только в факте, идут в конце.
    """
    mapping = match_people(fact_df["Сотрудник"].unique(), plan_people)
    fact_df = fact_df.assign(Сотрудник=fact_df["Сотрудник"].map(lambda n: mapping.get(n, n)))
    fact_by_person = fact_df.groupby("Сотрудник")[HOURS_COL].sum()

    rows = []
    for name, (df, norm, _) in plan_people.items():
        planned = round(float(df[label].sum()), 2) if label else 0.0
        self_fact = round(float(df[actual_label(label)].sum()), 2) if label else 0.0
        rows.append({"Сотрудник": name, "Норма, ч": norm, "План, ч": planned,
                     "Факт в планировщике, ч": self_fact,
                     "Факт по отметке, ч": round(float(fact_by_person.get(name, 0.0)), 2)})
    for name in fact_by_person.index:
        if name not in plan_people:
            rows.append({"Сотрудник": name, "Норма, ч": None, "План, ч": 0.0, "Факт в планировщике, ч": 0.0,
                         "Факт по отметке, ч": round(float(fact_by_person[name]), 2)})
    summary = pd.DataFrame(rows, columns=["Сотрудник", "Норма, ч", "План, ч", "Факт в планировщике, ч",
                                          "Факт по отметке, ч"])
    summary["Отклонение, ч"] = (summary["Факт по отметке, ч"] - summary["План, ч"]).round(2)
    summary["Факт / план"] = [f / p if p else None
                              for f, p in zip(summary["Факт по отметке, ч"], summary["План, ч"])]
    return summary, fact_df


def deviation_level(planned: float, fact: float) -> str:
    """ok — отклонение до 10% плана, warn — до 25%, bad — больше или план не заполнен при наличии факта."""
    if not planned:
        return "bad" if fact else "ok"
    dev = abs(fact - planned) / planned
    return "ok" if dev <= 0.10 else "warn" if dev <= 0.25 else "bad"


def fact_breakdown(fact_df: pd.DataFrame, by: str) -> pd.DataFrame:
    """Факт одного или нескольких сотрудников по выбранной колонке выгрузки: часы, доля, примечания."""
    if fact_df.empty:
        return pd.DataFrame(columns=[by, "Часы", "Доля", NOTE_COL])
    total = fact_df[HOURS_COL].sum()

    def notes(s):
        seen = []
        for n in s:
            if n and n not in seen:
                seen.append(n)
        return "; ".join(seen)

    out = (fact_df.groupby(by, sort=False)
           .agg(Часы=(HOURS_COL, "sum"), **{NOTE_COL: (NOTE_COL, notes)})
           .reset_index().sort_values("Часы", ascending=False))
    out["Часы"] = out["Часы"].round(2)
    out.insert(2, "Доля", out["Часы"] / total if total else 0.0)
    return out.reset_index(drop=True)


def build_comparison_workbook(summary: pd.DataFrame, fact_df: pd.DataFrame, by: str) -> bytes:
    """Excel: сводка план/факт, факт по сотрудникам и выбранной группировке, исходные строки факта."""
    by_person = (fact_df.groupby(["Сотрудник", by])[HOURS_COL].sum().round(2).reset_index()
                 .sort_values(["Сотрудник", HOURS_COL], ascending=[True, False]))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="План vs факт", index=False)
        by_person.to_excel(writer, sheet_name=f"Факт — {by}"[:31], index=False)
        fact_df.to_excel(writer, sheet_name="Отметки", index=False)
        for ws in writer.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)
        ws = writer.book["План vs факт"]
        for c in ws["G"][1:]:
            c.number_format = "0%"
    return buf.getvalue()
