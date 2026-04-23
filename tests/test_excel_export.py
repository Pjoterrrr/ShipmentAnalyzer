from __future__ import annotations

import ast
import io
import unittest
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from analytics_calendar import (
    build_calendar_frame,
    build_weekly_summary,
    classify_polish_day,
    get_last_completed_reference_week,
)


def load_export_functions() -> dict[str, object]:
    source_path = Path(__file__).resolve().parents[1] / "streamlit_app.py"
    source = source_path.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(source_path))
    required_names = {
        "blend_hex",
        "style_value",
        "format_date",
        "format_signed_int",
        "format_percent_display",
        "format_release_label",
        "format_release_summary",
        "get_reference_week_rows",
        "logo_available",
        "insert_logo",
        "style_excel_header",
        "autosize_worksheet",
        "decorate_delta_column",
        "excel_fill_color",
        "ensure_numeric_cells_black",
        "build_weekly_comparison_export",
        "build_report_matrix_export",
        "build_matrix_totals_export",
        "build_calendar_operational_export",
        "build_calendar_weekly_export",
        "add_weekly_comparison_chart",
        "add_totals_chart",
        "add_calendar_summary_chart",
        "style_table_region",
        "apply_number_formats",
        "style_multi_label_matrix_sheet",
        "apply_polish_calendar_highlights",
        "style_matrix_sheet",
        "highlight_calendar_rows",
        "highlight_weekly_rows",
        "write_summary_sheet",
        "to_excel_bytes",
    }
    selected_nodes = [
        node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in required_names
    ]
    module = ast.Module(body=selected_nodes, type_ignores=[])
    namespace: dict[str, object] = {
        "io": io,
        "pd": pd,
        "OpenpyxlImage": OpenpyxlImage,
        "Alignment": Alignment,
        "Border": Border,
        "BarChart": BarChart,
        "Font": Font,
        "LineChart": LineChart,
        "PatternFill": PatternFill,
        "Reference": Reference,
        "Side": Side,
        "get_column_letter": get_column_letter,
        "build_calendar_frame": build_calendar_frame,
        "classify_polish_day": classify_polish_day,
        "get_last_completed_reference_week": get_last_completed_reference_week,
        "BRAND_NAME": "Pjoter Development",
        "LOGO_PATH": Path("__missing_logo__.png"),
    }
    exec(compile(module, filename=str(source_path), mode="exec"), namespace)
    return namespace


def rgb_suffix(style_value) -> str:
    if style_value is None:
        return ""
    return str(style_value).upper()[-6:]


class ExcelExportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.export_functions = load_export_functions()

    def test_excel_export_contains_filtered_report_sheets_and_charts(self) -> None:
        detail_df = pd.DataFrame(
            [
                {
                    "PO Number": "PO-1",
                    "Part Number": "A1",
                    "Part Description": "Produkt A",
                    "Ship Date": pd.Timestamp("2026-05-01"),
                    "Receipt Date": pd.Timestamp("2026-05-04"),
                    "Unit of Measure": "HL",
                    "Quantity_Prev": 80.0,
                    "Quantity_Curr": 100.0,
                    "Delta": 20.0,
                    "Percent Change": 25.0,
                    "Change Direction": "Increase",
                    "Alert": True,
                    "Product Label": "A1 | Produkt A",
                },
                {
                    "PO Number": "PO-1",
                    "Part Number": "A1",
                    "Part Description": "Produkt A",
                    "Ship Date": pd.Timestamp("2026-05-02"),
                    "Receipt Date": pd.Timestamp("2026-05-05"),
                    "Unit of Measure": "HL",
                    "Quantity_Prev": 90.0,
                    "Quantity_Curr": 120.0,
                    "Delta": 30.0,
                    "Percent Change": 33.33,
                    "Change Direction": "Increase",
                    "Alert": True,
                    "Product Label": "A1 | Produkt A",
                },
            ]
        )
        weekly_summary = build_weekly_summary(
            detail_df,
            "Receipt Date",
            date(2026, 4, 27),
            date(2026, 5, 10),
            date(2026, 5, 10),
            15,
        )
        current_matrix_df = pd.DataFrame({"2026-05-04": [100.0], "2026-05-05": [120.0]}, index=["A1 | Produkt A"])
        delta_matrix_df = pd.DataFrame({"2026-05-04": [20.0], "2026-05-05": [30.0]}, index=["A1 | Produkt A"])
        prev_meta = {
            "po_number": "PO-1",
            "release_version": "15",
            "release_date": pd.Timestamp("2026-04-20"),
            "planner_name": "Planner",
            "planner_email": "planner@example.com",
        }
        curr_meta = {
            "po_number": "PO-1",
            "release_version": "16",
            "release_date": pd.Timestamp("2026-04-27"),
            "planner_name": "Planner",
            "planner_email": "planner@example.com",
        }
        product_summary = pd.DataFrame(
            [
                {
                    "Part Number": "A1",
                    "Part Description": "Produkt A",
                    "Product Label": "A1 | Produkt A",
                    "Quantity_Prev": 170.0,
                    "Quantity_Curr": 220.0,
                    "Delta": 50.0,
                    "Abs_Delta": 50.0,
                    "Alert_Count": 2,
                    "Change Direction": "Increase",
                }
            ]
        )

        to_excel_bytes = self.export_functions["to_excel_bytes"]
        excel_bytes = to_excel_bytes(
            detail_df,
            weekly_summary,
            current_matrix_df,
            delta_matrix_df,
            prev_meta,
            curr_meta,
            product_summary,
            "Receipt Date",
            date(2026, 4, 27),
            date(2026, 5, 10),
            [{"label": "Test", "title": "Produkt A", "copy": "Zmiana tygodniowa"}],
        )

        workbook = load_workbook(io.BytesIO(excel_bytes))
        self.assertIn("Weekly Comparison", workbook.sheetnames)
        self.assertIn("Calendar PL", workbook.sheetnames)
        self.assertIn("Current Matrix", workbook.sheetnames)
        self.assertIn("Delta Heatmap", workbook.sheetnames)

        calendar_sheet = workbook["Calendar PL"]
        calendar_rows = {
            calendar_sheet.cell(row=row, column=1).value: row
            for row in range(2, calendar_sheet.max_row + 1)
        }
        may_first_row = calendar_rows["2026-05-01"]
        may_third_row = calendar_rows["2026-05-03"]
        self.assertEqual(rgb_suffix(calendar_sheet.cell(row=may_first_row, column=1).fill.fgColor.rgb), "FEF3C7")
        self.assertEqual(rgb_suffix(calendar_sheet.cell(row=may_third_row, column=1).fill.fgColor.rgb), "FEF3C7")
        self.assertGreaterEqual(len(calendar_sheet._charts), 1)

        weekly_sheet = workbook["Weekly Comparison"]
        weekly_header = {cell.value: cell.column for cell in weekly_sheet[1]}
        current_release_cell = weekly_sheet.cell(row=2, column=weekly_header["Current Release Qty"])
        self.assertEqual(rgb_suffix(current_release_cell.font.color.rgb), "000000")
        self.assertGreaterEqual(len(weekly_sheet._charts), 1)

        current_matrix_sheet = workbook["Current Matrix"]
        delta_heatmap_sheet = workbook["Delta Heatmap"]
        self.assertGreaterEqual(len(current_matrix_sheet._charts), 1)
        self.assertGreaterEqual(len(delta_heatmap_sheet._charts), 1)


if __name__ == "__main__":
    unittest.main()
