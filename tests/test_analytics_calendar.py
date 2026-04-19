from __future__ import annotations

import unittest
from datetime import date

import pandas as pd

from analytics_calendar import (
    build_calendar_frame,
    build_weekly_summary,
    get_last_completed_reference_week,
    week_label_for_date,
)


class AnalyticsCalendarTests(unittest.TestCase):
    def test_reference_week_uses_sunday_as_completed_week_end(self) -> None:
        reference = get_last_completed_reference_week(date(2026, 4, 19))
        self.assertEqual(reference.week_label, "2026-W16")
        self.assertEqual(reference.week_start, date(2026, 4, 13))
        self.assertEqual(reference.week_end, date(2026, 4, 19))

    def test_reference_week_uses_previous_week_for_midweek_reference(self) -> None:
        reference = get_last_completed_reference_week(date(2026, 4, 16))
        self.assertEqual(reference.week_label, "2026-W15")
        self.assertEqual(reference.week_start, date(2026, 4, 6))
        self.assertEqual(reference.week_end, date(2026, 4, 12))

    def test_iso_week_label_is_correct_on_year_boundary(self) -> None:
        self.assertEqual(week_label_for_date(date(2021, 1, 1)), "2020-W53")

    def test_polish_calendar_excludes_weekends_and_public_holidays(self) -> None:
        calendar = build_calendar_frame(date(2026, 4, 27), date(2026, 5, 3))
        working_days = int(calendar["Is Working Day"].sum())
        may_first = calendar.loc[calendar["Date"].eq(pd.Timestamp("2026-05-01"))].iloc[0]
        self.assertEqual(working_days, 4)
        self.assertTrue(bool(may_first["Is Holiday"]))
        self.assertEqual(may_first["Holiday Name"], "Swieto Pracy")

    def test_weekly_summary_reports_release_and_wow_changes(self) -> None:
        frame = pd.DataFrame(
            [
                {
                    "Receipt Date": pd.Timestamp("2026-04-28"),
                    "Quantity_Prev": 80,
                    "Quantity_Curr": 100,
                    "Delta": 20,
                    "Alert": True,
                    "Product Label": "A | Produkt A",
                },
                {
                    "Receipt Date": pd.Timestamp("2026-05-05"),
                    "Quantity_Prev": 90,
                    "Quantity_Curr": 120,
                    "Delta": 30,
                    "Alert": True,
                    "Product Label": "A | Produkt A",
                },
            ]
        )

        weekly = build_weekly_summary(
            frame,
            "Receipt Date",
            date(2026, 4, 27),
            date(2026, 5, 10),
            date(2026, 5, 10),
            15,
        )

        week_18 = weekly.loc[weekly["Week Label"].eq("2026-W18")].iloc[0]
        week_19 = weekly.loc[weekly["Week Label"].eq("2026-W19")].iloc[0]

        self.assertEqual(float(week_18["Quantity_Curr"]), 100.0)
        self.assertEqual(float(week_19["Quantity_Curr"]), 120.0)
        self.assertEqual(week_19["Release Percent Label"], "+33.33%")
        self.assertEqual(week_19["WoW Percent Label"], "+20.00%")
        self.assertEqual(int(week_18["Working_Days_PL"]), 4)
        self.assertEqual(int(week_19["Working_Days_PL"]), 5)
        self.assertTrue(bool(week_19["Is Reference Week"]))


if __name__ == "__main__":
    unittest.main()
