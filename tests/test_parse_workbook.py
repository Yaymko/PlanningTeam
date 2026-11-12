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
    assert ticket_row["Статус"] == "в работе"
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


def test_parses_sheets_without_title_row_and_short_totals_label():
    """Выгрузка, где заголовок стоит в первой строке, а итоговая строка подписана просто «Итого»."""
    labels = ["05.10–11.10 ч", "12.10–18.10 ч"]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Петров Пётр Петрович")
    ws.append(["Тикет", "Название", "Колонка", "Доска", "Срок", "В работе", *labels, "Итого, ч."])
    ws.append(["ABC-2", "Другая штука", "10 Работа", "1743", None, True, 3, None, None])
    ws.append([None, None, "Доп. работа", None, None, None, None, None, None])
    ws.append(["Итого", None, None, None, None, None, "=SUM(G2:G3)", "=SUM(H2:H3)", None])
    buf = io.BytesIO()
    wb.save(buf)

    people_data, week_labels = parse_workbook(buf.getvalue())

    assert week_labels == labels
    df = people_data["Петров Пётр Петрович"]
    assert list(df["Тикет"]) == ["ABC-2", ""]
    assert df.iloc[0][labels[0]] == 3.0


def test_parses_status_and_priority_columns():
    """Выгрузка, где уже есть колонки «Статус» и «Приоритет», — недели идут после них."""
    labels = ["05.10–11.10 ч", "12.10–18.10 ч"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Петров Пётр Петрович"
    ws.append(["Тикет", "Название", "Колонка", "Доска", "Срок", "Статус", "Приоритет", *labels, "Итого, ч."])
    ws.append(["ABC-2", "Другая штука", "10 Работа", "1743", None, "пауза", "срочно", 3, 1, None])
    buf = io.BytesIO()
    wb.save(buf)

    people_data, week_labels = parse_workbook(buf.getvalue())

    assert week_labels == labels
    row = people_data["Петров Пётр Петрович"].iloc[0]
    assert (row["Статус"], row["Приоритет"], row[labels[0]], row[labels[1]]) == ("пауза", "срочно", 3.0, 1.0)


def test_legacy_in_work_flag_becomes_status():
    people_data, _ = parse_workbook(_build_sample_workbook())
    assert people_data["Иванов Иван Иванович"]["Статус"].tolist() == ["в работе", ""]


def test_new_overview_sheet_names_and_assignees_column_are_ignored():
    """Новая выгрузка: обзорные листы называются «Обзор 1267» и содержат колонку «Исполнители»."""
    overview_headers = ["Тикет", "Название", "Колонка", "Исполнители", "Срок", "В работе", *WEEK_LABELS]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title in ["Обзор 1267", "Сводка доски"]:
        ws = wb.create_sheet(title)
        ws.append([f"{title} — справочно"])
        ws.append(overview_headers)
        ws.append(["ABC-1", "Сделать штуку", "10 Работа", "Иванов Иван Иванович", None, True, None, None])
    ws = wb.create_sheet("Обзоров О.")
    ws.append(["Обзоров Олег Олегович"])
    ws.append(HEADERS)
    ws.append(["ABC-1", "Сделать штуку", "10 Работа", "1267", None, True, 2, 0, None])
    buf = io.BytesIO()
    wb.save(buf)

    people_data, week_labels = parse_workbook(buf.getvalue())

    # Полное ФИО из строки над шапкой того же человека, что и в названии листа.
    assert list(people_data.keys()) == ["Обзоров Олег Олегович"]
    assert week_labels == WEEK_LABELS


def _build_new_format_workbook(hours=None) -> bytes:
    """Выгрузка с ноября 2026: одна неделя «16.11–22.11, ч», без колонок «В работе» и «Итого, ч.»,
    лист назван коротко, полное ФИО — в строке над шапкой, итоговая строка «Итого, ч.»."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ov = wb.create_sheet("Обзор 1267")
    ov.append(["Доска 1267 — статус на 12.11.2026"])
    ov.append(["Тикет", "Название", "Колонка", "Исполнители", "Срок"])
    ov.append(["ABC-1", "Сделать штуку", "10 Работа", "Панькин Александр Викторович", None])
    ws = wb.create_sheet("Панькин А. В.")
    ws.append(["Панькин Александр Викторович"])
    ws.append(["Тикет", "Название", "Колонка", "Доска", "Срок", "16.11–22.11, ч"])
    ws.append(["ABC-1", "Сделать штуку", "10 Работа", "1267", None, hours])
    ws.append([None, "Доп. работа", "Доп. работа", None, None, None])
    ws.append(["Итого, ч.", None, None, None, None, "=SUM(F3:F4)"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_parses_new_single_week_format():
    people_data, week_labels = parse_workbook(_build_new_format_workbook(hours=3))

    assert week_labels == ["16.11–22.11 ч"]  # «, ч» приводится к старому виду «… ч»
    assert list(people_data) == ["Панькин Александр Викторович"]
    df = people_data["Панькин Александр Викторович"]
    assert list(df["Тикет"]) == ["ABC-1", ""]
    assert df.loc[0, "16.11–22.11 ч"] == 3.0
    assert df.loc[1, "Название"] == "Доп. работа"
    assert df.loc[0, "Статус"] != "в работе"  # колонки «В работе» нет


def test_week_label_helpers_accept_comma_suffix():
    from core import actual_label, canonical_week_label, clean_week_label, is_week_header

    assert canonical_week_label("16.11–22.11, ч") == "16.11–22.11 ч"
    assert canonical_week_label("28.09–04.10 ч") == "28.09–04.10 ч"
    assert clean_week_label("16.11–22.11, ч") == "16.11–22.11"
    assert actual_label("16.11–22.11, ч") == "16.11–22.11 факт, ч"
    assert is_week_header("16.11–22.11, ч") and is_week_header("28.09–04.10 ч")
    assert not is_week_header("Итого, ч.") and not is_week_header("Срок")


def test_new_format_upload_keeps_data_saved_under_short_name(tmp_path):
    """Часы, норма и доп. работы, сохранённые под коротким именем, переходят к полному ФИО."""
    from datetime import date

    import pandas as pd

    from storage import Store

    week = "16.11–22.11 ч"
    store = Store(tmp_path / "db.sqlite")
    first, _ = parse_workbook(_build_new_format_workbook())
    short = {"Панькин А.": first["Панькин Александр Викторович"]}
    store.save_plan("old.xlsx", short, [week], today=date(2026, 11, 12))
    df, _, _ = store.load_person("Панькин А.")
    df.loc[0, week] = 5.0
    extra = {c: "" for c in df.columns} | {"Колонка": "Доп. работа", "Название": "Созвоны", "Срок": None,
                                            week: 2.0, "16.11–22.11 факт, ч": 0.0}
    store.save_person("Панькин А.", pd.concat([df, pd.DataFrame([extra])], ignore_index=True), 32)

    people_data, week_labels = parse_workbook(_build_new_format_workbook())
    store.save_plan("new.xlsx", people_data, week_labels, today=date(2026, 11, 12))

    df, norm, _ = store.load_person("Панькин Александр Викторович")
    assert norm == 32.0
    assert df.loc[df["Тикет"] == "ABC-1", week].tolist() == [5.0]
    assert df.loc[df["Название"] == "Созвоны", week].tolist() == [2.0]
