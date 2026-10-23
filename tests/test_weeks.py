"""Only the nearest weeks of the plan are shown: the current one and the two after it."""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import nearest_weeks  # noqa: E402

WEEKS = ["28.09–04.10 ч", "05.10–11.10 ч", "12.10–18.10 ч", "19.10–25.10 ч", "26.10–01.11 ч"]


def test_current_and_two_next_weeks():
    assert nearest_weeks(WEEKS, date(2026, 10, 5)) == WEEKS[1:4]  # Monday: current, next, the week after
    assert nearest_weeks(WEEKS, date(2026, 10, 2)) == WEEKS[0:3]
    assert nearest_weeks(WEEKS, date(2026, 10, 14)) == WEEKS[2:5]


def test_near_the_end_and_outside_the_plan():
    assert nearest_weeks(WEEKS, date(2026, 10, 23)) == WEEKS[2:5]  # fewer than 3 weeks left: the last 3
    assert nearest_weeks(WEEKS, date(2026, 9, 20)) == WEEKS[0:3]
    assert nearest_weeks(WEEKS, date(2026, 11, 10)) == WEEKS[2:5]


def test_week_across_new_year():
    weeks = ["22.12–28.12 ч", "29.12–04.01 ч", "05.01–11.01 ч", "12.01–18.01 ч"]
    assert nearest_weeks(weeks, date(2027, 1, 2)) == weeks[1:4]
    assert nearest_weeks(weeks, date(2026, 12, 23)) == weeks[0:3]


def test_short_plan_and_unreadable_labels_show_everything():
    assert nearest_weeks(WEEKS[:2], date(2026, 10, 5)) == WEEKS[:2]
    assert nearest_weeks(["Неделя 1", "Неделя 2", "Неделя 3", "Неделя 4"], date(2026, 10, 5)) == [
        "Неделя 1", "Неделя 2", "Неделя 3", "Неделя 4"]
