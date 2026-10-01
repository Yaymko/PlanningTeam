"""Unit test for the board-export parsing logic in app.py."""
import io
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import parse_workbook  # noqa: E402

WEEK_LABELS = ["28.09–04.10 ч", "05.10–11.10 ч"]
HEADERS = ["Тикет", "Название", "Колонка", "Доска", "Срок", "В работе", *WEEK_LABELS, "Итого, ч."]


def _build_sample_workbook() -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Иванов Иван Иванович")
    ws.append(["Иванов Иван Иванович — тестовые данные"])
    ws.append(HEADERS)
    ws.append(["ABC-1", "Сделать штуку", "10 Работа", "1267", None, True, 4, 0, None])
    ws.append([None, None, "Доп. работа", None, None, None, None, None, None])
    ws.append(["Итого по неделям:", None, None, None, None, None, "=SUM(G3:G4)", "=SUM(H3:H4)", None])

    overview = wb.create_sheet("Обзор тикетов 1267")
    overview.append(["Сводный отчёт — справочно, не парсится"])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_parses_week_labels_and_rows():
    people_data, week_labels = parse_workbook(_build_sample_workbook())

    assert week_labels == WEEK_LABELS
    assert list(people_data.keys()) == ["Иванов Иван Иванович"]

    df = people_data["Иванов Иван Иванович"]
    assert len(df) == 2  # one ticket row + one "Доп. работа" filler row, stops before totals row

    ticket_row = df.iloc[0]
    assert ticket_row["Тикет"] == "ABC-1"
    assert bool(ticket_row["В работе"]) is True
    assert ticket_row[WEEK_LABELS[0]] == 4.0
    assert ticket_row[WEEK_LABELS[1]] == 0.0

    filler_row = df.iloc[1]
    assert filler_row["Колонка"] == "Доп. работа"


def test_overview_sheets_are_ignored():
    people_data, _ = parse_workbook(_build_sample_workbook())
    assert "Обзор тикетов 1267" not in people_data


def test_missing_week_columns_returns_empty_labels():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Обзор тикетов 1267")
    ws.append(["Только справочный лист, без исполнителей"])
    buf = io.BytesIO()
    wb.save(buf)

    people_data, week_labels = parse_workbook(buf.getvalue())
    assert week_labels == []
    assert people_data == {}
