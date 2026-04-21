from __future__ import annotations

import io
import unittest
from datetime import date

import pandas as pd
from openpyxl import Workbook

from analytics_calendar import build_weekly_summary
from release_loader import compare_releases, detect_file_type, load_release


def build_legacy_workbook_from_rows(
    rows: list[dict[str, object]],
    raw_sheet_name: str = "Raw",
    include_overview: bool = True,
) -> bytes:
    output = io.BytesIO()
    overview_df = pd.DataFrame(
        [
            {
                "Planner Name": "Legacy Planner",
                "Planner Email": "legacy@example.com",
            }
        ]
    )
    raw_df = pd.DataFrame(rows)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if include_overview:
            overview_df.to_excel(writer, sheet_name="Overview", startrow=4, index=False)
        raw_df.to_excel(writer, sheet_name=raw_sheet_name, index=False)

    return output.getvalue()


def build_legacy_workbook(
    open_quantity: float = 120.0,
    raw_sheet_name: str = "Raw",
    include_overview: bool = True,
) -> bytes:
    return build_legacy_workbook_from_rows(
        [
            {
                "PO Number": "PO-LEG-1",
                "PO Line #": "10",
                "Release Version": "15",
                "Release Date": pd.Timestamp("2026-04-18"),
                "Part Number": "LEG-001",
                "Part Description": "Legacy Product",
                "Ship Date": pd.Timestamp("2026-04-21"),
                "Receipt Date": pd.Timestamp("2026-04-24"),
                "Open Quantity": open_quantity,
                "Unit of Measure": "EA",
            }
        ],
        raw_sheet_name=raw_sheet_name,
        include_overview=include_overview,
    )


def build_vl10e_workbook(
    first_qty: float = 80.0,
    include_removed_row: bool = True,
    include_new_block: bool = False,
) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "VL10E"
    worksheet["B1"] = "OriginDoc."
    worksheet["C1"] = "Item"
    worksheet["D1"] = "Ship-to"
    worksheet["F1"] = "Material"
    worksheet["H1"] = "Description"
    worksheet["J1"] = "Unrestr."
    worksheet["K1"] = "Unl. Point"
    worksheet["L1"] = "Cust.mat."
    worksheet["M1"] = "PO Number"

    worksheet["B2"] = 500001
    worksheet["C2"] = 10
    worksheet["D2"] = "SHIP-01"
    worksheet["F2"] = "MAT-001"
    worksheet["H2"] = "Mercury"
    worksheet["J2"] = 150
    worksheet["K2"] = "PL01"
    worksheet["L2"] = "CM-001"
    worksheet["M2"] = "PO-VL-1"

    worksheet["B3"] = date(2026, 4, 21)
    worksheet["C3"] = date(2026, 4, 22)
    worksheet["E3"] = first_qty
    worksheet["G3"] = "EA"
    worksheet["H3"] = first_qty
    worksheet["I3"] = "EA"

    if include_removed_row:
        worksheet["B4"] = date(2026, 4, 28)
        worksheet["C4"] = date(2026, 4, 29)
        worksheet["E4"] = 40
        worksheet["G4"] = "EA"
        worksheet["H4"] = first_qty + 40
        worksheet["I4"] = "EA"

    worksheet["B5"] = 500002
    worksheet["C5"] = 20
    worksheet["D5"] = "SHIP-02"
    worksheet["F5"] = "MAT-002"
    worksheet["H5"] = "Venus"
    worksheet["M5"] = "PO-VL-2"

    worksheet["B6"] = date(2026, 4, 23)
    worksheet["C6"] = date(2026, 4, 24)
    worksheet["E6"] = 10
    worksheet["G6"] = "EA"
    worksheet["H6"] = 10
    worksheet["I6"] = "EA"

    if include_new_block:
        worksheet["B8"] = 500003
        worksheet["C8"] = 30
        worksheet["D8"] = "SHIP-03"
        worksheet["F8"] = "MAT-003"
        worksheet["H8"] = "Mars"
        worksheet["M8"] = "PO-VL-3"

        worksheet["B9"] = date(2026, 4, 25)
        worksheet["C9"] = date(2026, 4, 26)
        worksheet["E9"] = 30
        worksheet["G9"] = "EA"
        worksheet["H9"] = 30
        worksheet["I9"] = "EA"

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def build_weekly_pivot_workbook() -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet2"

    worksheet["C1"] = 2026
    worksheet["F1"] = 2027

    worksheet["A2"] = "Row Labels"
    worksheet["C2"] = "backlog"
    worksheet["D2"] = 17
    worksheet["E2"] = 18
    worksheet["F2"] = 1

    worksheet["A3"] = "MAT-001"
    worksheet["C3"] = 5
    worksheet["D3"] = 100
    worksheet["E3"] = 50
    worksheet["F3"] = 70

    worksheet["A4"] = "MAT-002"
    worksheet["C4"] = 0
    worksheet["D4"] = 30
    worksheet["E4"] = 10
    worksheet["F4"] = 20

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


class ReleaseLoaderTests(unittest.TestCase):
    def test_detect_file_type_recognizes_all_formats(self) -> None:
        legacy_bytes = build_legacy_workbook()
        vl10e_bytes = build_vl10e_workbook()
        weekly_bytes = build_weekly_pivot_workbook()

        self.assertEqual(detect_file_type(legacy_bytes), "legacy_wide")
        self.assertEqual(detect_file_type(vl10e_bytes), "vl10e_block")
        self.assertEqual(detect_file_type(weekly_bytes), "cw_weekly_pivot")

    def test_detect_file_type_uses_first_sheet_when_raw_is_missing(self) -> None:
        legacy_without_raw = build_legacy_workbook(
            open_quantity=140.0,
            raw_sheet_name="Data",
            include_overview=False,
        )

        self.assertEqual(detect_file_type(legacy_without_raw), "legacy_wide")

    def test_load_release_supports_legacy_format(self) -> None:
        legacy_bytes = build_legacy_workbook(open_quantity=125.0)

        dataframe, metadata = load_release(legacy_bytes, "legacy release 2026-04-18.xlsx")

        self.assertEqual(metadata["file_type"], "legacy_wide")
        self.assertEqual(metadata["po_number"], "PO-LEG-1")
        self.assertEqual(metadata["release_version"], "15")
        self.assertEqual(metadata["planner_name"], "Legacy Planner")
        self.assertIn("Part Number", dataframe.columns)
        self.assertIn("Ship Date", dataframe.columns)
        self.assertEqual(float(dataframe.iloc[0]["Open Quantity"]), 125.0)

    def test_load_release_supports_legacy_format_without_raw_sheet(self) -> None:
        legacy_bytes = build_legacy_workbook(
            open_quantity=140.0,
            raw_sheet_name="Data",
            include_overview=False,
        )

        dataframe, metadata = load_release(legacy_bytes, "legacy fallback 2026-04-18.xlsx")

        self.assertEqual(metadata["file_type"], "legacy_wide")
        self.assertEqual(metadata["sheet_name"], "Data")
        self.assertEqual(metadata["sheet_names"], ["Data"])
        self.assertEqual(metadata["planner_name"], "n/a")
        self.assertEqual(float(dataframe.iloc[0]["Open Quantity"]), 140.0)

    def test_load_release_supports_vl10e_block_format(self) -> None:
        vl10e_bytes = build_vl10e_workbook(first_qty=80.0)

        dataframe, metadata = load_release(vl10e_bytes, "VL10E Merc.xls 20.04.2026.xlsx")

        self.assertEqual(metadata["file_type"], "vl10e_block")
        self.assertEqual(metadata["sheet_name"], "VL10E")
        self.assertEqual(metadata["release_version"], "2026-04-20")
        self.assertEqual(str(pd.Timestamp(metadata["release_date"]).date()), "2026-04-20")
        self.assertEqual(dataframe.iloc[0]["Origin Doc"], "500001")
        self.assertEqual(dataframe.iloc[0]["Item"], "10")
        self.assertEqual(dataframe.iloc[0]["Part Number"], "MAT-001")
        self.assertEqual(dataframe.iloc[0]["PO Number"], "PO-VL-1")

    def test_load_release_supports_weekly_pivot_format(self) -> None:
        weekly_bytes = build_weekly_pivot_workbook()

        dataframe, metadata = load_release(weekly_bytes, "CW17_Megatech Q7Q9.xlsx")

        self.assertEqual(metadata["file_type"], "cw_weekly_pivot")
        self.assertEqual(metadata["sheet_name"], "Sheet2")
        self.assertEqual(metadata["release_version"], "2026-04-20")
        self.assertEqual(str(pd.Timestamp(metadata["release_date"]).date()), "2026-04-20")

        mat001 = dataframe[dataframe["Part Number"] == "MAT-001"].reset_index(drop=True)
        self.assertEqual(mat001["Week Label"].tolist(), ["2026-W17", "2026-W18", "2027-W01"])
        self.assertEqual(mat001["ISO Year"].tolist(), [2026, 2026, 2027])
        self.assertEqual(mat001["ISO Week"].tolist(), [17, 18, 1])
        self.assertEqual(mat001["Open Quantity"].tolist(), [100.0, 50.0, 70.0])
        self.assertEqual(mat001["Backlog"].tolist(), [5.0, 5.0, 5.0])
        self.assertTrue((mat001["Time Bucket"] == "weekly").all())
        self.assertEqual(
            mat001["Receipt Date"].dt.strftime("%Y-%m-%d").tolist(),
            ["2026-04-20", "2026-04-27", "2027-01-04"],
        )

    def test_load_release_returns_readable_error_for_invalid_file(self) -> None:
        with self.assertRaisesRegex(ValueError, "Nie udało się odczytać pliku Excel"):
            load_release(b"not an excel file", "broken.xlsx")

    def test_vl10e_comparison_marks_new_and_removed_demand_and_supports_weekly_rollup(self) -> None:
        prev_bytes = build_vl10e_workbook(first_qty=80.0, include_removed_row=True, include_new_block=False)
        curr_bytes = build_vl10e_workbook(first_qty=100.0, include_removed_row=False, include_new_block=True)

        prev_df, _ = load_release(prev_bytes, "VL10E Merc.xls 13.04.2026.xlsx")
        curr_df, _ = load_release(curr_bytes, "VL10E Merc.xls 20.04.2026.xlsx")
        result = compare_releases(prev_df, curr_df, threshold=15)

        increased_row = result[
            (result["Part Number"] == "MAT-001")
            & (result["Receipt Date"] == pd.Timestamp("2026-04-22"))
        ].iloc[0]
        removed_row = result[
            (result["Part Number"] == "MAT-001")
            & (result["Receipt Date"] == pd.Timestamp("2026-04-29"))
        ].iloc[0]
        new_row = result[
            (result["Part Number"] == "MAT-003")
            & (result["Receipt Date"] == pd.Timestamp("2026-04-26"))
        ].iloc[0]

        self.assertEqual(float(increased_row["Quantity_Prev"]), 80.0)
        self.assertEqual(float(increased_row["Quantity_Curr"]), 100.0)
        self.assertEqual(float(increased_row["Percent Change"]), 25.0)
        self.assertTrue(bool(increased_row["Alert"]))

        self.assertEqual(removed_row["Demand Status"], "REMOVED DEMAND")
        self.assertEqual(float(removed_row["Quantity_Curr"]), 0.0)
        self.assertEqual(float(removed_row["Percent Change"]), -100.0)

        self.assertEqual(new_row["Demand Status"], "NEW DEMAND")
        self.assertEqual(float(new_row["Quantity_Prev"]), 0.0)
        self.assertEqual(float(new_row["Quantity_Curr"]), 30.0)
        self.assertEqual(float(new_row["Percent Change"]), 100.0)

        weekly_summary = build_weekly_summary(
            result,
            "Receipt Date",
            date(2026, 4, 20),
            date(2026, 4, 29),
            date(2026, 4, 29),
            15,
        )

        week_17 = weekly_summary[weekly_summary["Week Label"] == "2026-W17"].iloc[0]
        week_18 = weekly_summary[weekly_summary["Week Label"] == "2026-W18"].iloc[0]

        self.assertEqual(float(week_17["Quantity_Prev"]), 90.0)
        self.assertEqual(float(week_17["Quantity_Curr"]), 140.0)
        self.assertEqual(float(week_18["Quantity_Prev"]), 40.0)
        self.assertEqual(float(week_18["Quantity_Curr"]), 0.0)

    def test_compare_releases_rolls_daily_data_up_to_week_when_weekly_file_is_used(self) -> None:
        legacy_rows = [
            {
                "PO Number": "PO-LEG-1",
                "PO Line #": "10",
                "Release Version": "15",
                "Release Date": pd.Timestamp("2026-04-18"),
                "Part Number": "MAT-001",
                "Part Description": "Daily Product",
                "Ship Date": pd.Timestamp("2026-04-20"),
                "Receipt Date": pd.Timestamp("2026-04-21"),
                "Open Quantity": 30.0,
                "Unit of Measure": "EA",
            },
            {
                "PO Number": "PO-LEG-2",
                "PO Line #": "20",
                "Release Version": "15",
                "Release Date": pd.Timestamp("2026-04-18"),
                "Part Number": "MAT-001",
                "Part Description": "Daily Product",
                "Ship Date": pd.Timestamp("2026-04-21"),
                "Receipt Date": pd.Timestamp("2026-04-22"),
                "Open Quantity": 70.0,
                "Unit of Measure": "EA",
            },
        ]
        legacy_bytes = build_legacy_workbook_from_rows(legacy_rows)
        weekly_bytes = build_weekly_pivot_workbook()

        prev_df, _ = load_release(legacy_bytes, "legacy release 2026-04-18.xlsx")
        curr_df, _ = load_release(weekly_bytes, "CW17_Megatech Q7Q9.xlsx")
        result = compare_releases(prev_df, curr_df, threshold=15)

        mat001_week17 = result[
            (result["Part Number"] == "MAT-001")
            & (result["Week Label"] == "2026-W17")
        ].iloc[0]

        self.assertEqual(float(mat001_week17["Quantity_Prev"]), 100.0)
        self.assertEqual(float(mat001_week17["Quantity_Curr"]), 100.0)
        self.assertEqual(float(mat001_week17["Delta"]), 0.0)
        self.assertEqual(mat001_week17["Demand Status"], "")
        self.assertEqual(str(pd.Timestamp(mat001_week17["Receipt Date"]).date()), "2026-04-20")


if __name__ == "__main__":
    unittest.main()
