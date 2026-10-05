"""Only the nearest weeks of the plan are shown by default."""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import nearest_weeks  # noqa: E402

WEEKS = ["28.09–04.10 ч", "05.10–11.10 ч", "12.10–18.10 ч", "19.10–25.10 ч"]


def test_current_and_next_week():
    assert nearest_weeks(WEEKS, date(2026, 10, 7)) == WEEKS[1:3]
    assert nearest_weeks(WEEKS, date(2026, 10, 4)) == WEEKS[0:2]  # last day of the first week


def test_before_and_after_the_plan():
    assert nearest_weeks(WEEKS, date(2026, 9, 20)) == WEEKS[0:2]
    assert nearest_weeks(WEEKS, date(2026, 10, 23)) == WEEKS[2:4]
    assert nearest_weeks(WEEKS, date(2026, 11, 10)) == WEEKS[2:4]


def test_week_across_new_year():
    weeks = ["22.12–28.12 ч", "29.12–04.01 ч", "05.01–11.01 ч"]
    assert nearest_weeks(weeks, date(2027, 1, 2)) == weeks[1:3]
    assert nearest_weeks(weeks, date(2026, 12, 24)) == weeks[0:2]


def test_unreadable_labels_show_everything():
    assert nearest_weeks(["Неделя 1", "Неделя 2", "Неделя 3"], date(2026, 10, 5)) == [
        "Неделя 1", "Неделя 2", "Неделя 3"]
