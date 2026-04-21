from __future__ import annotations

import unittest
from datetime import date

import pandas as pd

from planner_helpers import (
    build_planner_input_frame,
    calculate_planner_outputs,
    planner_inputs_to_state,
    prepare_planner_source,
)


class PlannerHelpersTests(unittest.TestCase):
    def test_prepare_planner_source_groups_positive_current_demand_by_ship_date(self) -> None:
        frame = pd.DataFrame(
            [
                {
                    "Part Number": "A-100",
                    "Part Description": "Produkt A",
                    "Ship Date": pd.Timestamp("2026-04-25"),
                    "Quantity_Curr": 10,
                },
                {
                    "Part Number": "A-100",
                    "Part Description": "Produkt A",
                    "Ship Date": pd.Timestamp("2026-04-25"),
                    "Quantity_Curr": 5,
                },
                {
                    "Part Number": "A-100",
                    "Part Description": "Produkt A",
                    "Ship Date": pd.Timestamp("2026-04-26"),
                    "Quantity_Curr": 0,
                },
            ]
        )

        planner_source = prepare_planner_source(frame)

        self.assertEqual(len(planner_source), 1)
        self.assertEqual(planner_source.iloc[0]["Part Number"], "A-100")
        self.assertEqual(float(planner_source.iloc[0]["Demand Qty"]), 15.0)

    def test_calculate_planner_outputs_returns_shortage_priority_and_daily_detail(self) -> None:
        planner_source = prepare_planner_source(
            pd.DataFrame(
                [
                    {
                        "Part Number": "A-100",
                        "Part Description": "Produkt A",
                        "Ship Date": pd.Timestamp("2026-04-22"),
                        "Quantity_Curr": 20,
                    },
                    {
                        "Part Number": "A-100",
                        "Part Description": "Produkt A",
                        "Ship Date": pd.Timestamp("2026-04-23"),
                        "Quantity_Curr": 20,
                    },
                    {
                        "Part Number": "B-200",
                        "Part Description": "Produkt B",
                        "Ship Date": pd.Timestamp("2026-04-28"),
                        "Quantity_Curr": 15,
                    },
                ]
            )
        )
        planner_inputs = build_planner_input_frame(
            planner_source,
            planner_inputs_to_state(
                pd.DataFrame(
                    [
                        {
                            "Part Number": "A-100",
                            "Part Description": "Produkt A",
                            "Stock": 25,
                            "Safety Stock": 5,
                        },
                        {
                            "Part Number": "B-200",
                            "Part Description": "Produkt B",
                            "Stock": 40,
                            "Safety Stock": 5,
                        },
                    ]
                )
            ),
        )

        planner_results, planner_daily = calculate_planner_outputs(
            planner_source,
            planner_inputs,
            today=date(2026, 4, 21),
        )

        first_row = planner_results.iloc[0]
        self.assertEqual(first_row["Part Number"], "A-100")
        self.assertEqual(first_row["Status"], "Krytyczne")
        self.assertEqual(float(first_row["Shortage Qty"]), 20.0)
        self.assertEqual(int(first_row["Production Priority"]), 1)

        covered_row = planner_results.loc[planner_results["Part Number"].eq("B-200")].iloc[0]
        self.assertEqual(covered_row["Status"], "Pokryte")
        self.assertEqual(float(covered_row["Qty To Produce Now"]), 0.0)

        detail_rows = planner_daily.loc[planner_daily["Part Number"].eq("A-100")]
        self.assertEqual(len(detail_rows), 2)
        self.assertTrue(bool(detail_rows.iloc[-1]["Shortage Flag"]))


if __name__ == "__main__":
    unittest.main()
