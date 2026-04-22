import io

import altair as alt
import pandas as pd
import streamlit as st


REPORT_VIEWS = {
    "executive": "Executive Report",
    "weekly_risk": "Weekly Risk Report",
    "product_change": "Product Change Report",
    "alert": "Alert Report",
    "demand_change": "Demand Change Report",
    "top_risk": "Top Risk Parts",
}


def _interactive_chart(chart):
    if chart is None:
        return None
    try:
        return chart.interactive()
    except Exception:
        return chart


def _make_excel_bytes(report_tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, dataframe in report_tables.items():
            safe_sheet = str(sheet_name)[:31]
            export_df = dataframe.copy()
            export_df.to_excel(writer, sheet_name=safe_sheet, index=False)
    return output.getvalue()


def _format_report_table(dataframe):
    if dataframe is None or dataframe.empty:
        return pd.DataFrame()

    table = dataframe.copy()
    for column in table.columns:
        if pd.api.types.is_datetime64_any_dtype(table[column]):
            table[column] = pd.to_datetime(table[column], errors="coerce").dt.strftime("%Y-%m-%d")
    return table


def _build_executive_report(filtered_df, product_summary, date_summary, weekly_summary, reference):
    total_prev = float(filtered_df["Quantity_Prev"].sum())
    total_curr = float(filtered_df["Quantity_Curr"].sum())
    total_delta = float(filtered_df["Delta"].sum())
    alert_rows = int(filtered_df["Alert"].sum()) if "Alert" in filtered_df.columns else 0
    changed_products = int((product_summary["Delta"] != 0).sum()) if not product_summary.empty else 0

    largest_increase = "n/a"
    if not product_summary.empty and (product_summary["Delta"] > 0).any():
        top_increase = product_summary.nlargest(1, "Delta").iloc[0]
        largest_increase = f"{top_increase['Part Number']} | {top_increase['Delta']:+,.0f}"

    largest_drop = "n/a"
    if not product_summary.empty and (product_summary["Delta"] < 0).any():
        top_drop = product_summary.nsmallest(1, "Delta").iloc[0]
        largest_drop = f"{top_drop['Part Number']} | {top_drop['Delta']:+,.0f}"

    peak_alert_day = "n/a"
    if not date_summary.empty:
        peak_row = date_summary.sort_values(["Alerts", "Delta"], ascending=[False, False]).iloc[0]
        peak_alert_day = pd.Timestamp(peak_row["Analysis Date"]).strftime("%Y-%m-%d")

    weekly_alerts = int(weekly_summary["Any Weekly Alert"].sum()) if "Any Weekly Alert" in weekly_summary.columns else 0

    return pd.DataFrame(
        [
            {"Metric": "Total Previous Qty", "Value": f"{total_prev:,.0f}", "Commentary": "Suma poprzedniego release'u."},
            {"Metric": "Total Current Qty", "Value": f"{total_curr:,.0f}", "Commentary": "Suma aktualnego release'u."},
            {"Metric": "Net Delta", "Value": f"{total_delta:+,.0f}", "Commentary": "Bilans zmian w aktywnym zakresie."},
            {"Metric": "Alert Rows", "Value": f"{alert_rows:,}", "Commentary": "Liczba wierszy przekraczajacych prog alertu."},
            {"Metric": "Products Changed", "Value": f"{changed_products:,}", "Commentary": "Produkty ze zmiana wolumenu."},
            {"Metric": "Largest Increase", "Value": largest_increase, "Commentary": "Najwiekszy wzrost po produkcie."},
            {"Metric": "Largest Drop", "Value": largest_drop, "Commentary": "Najwiekszy spadek po produkcie."},
            {"Metric": "Peak Alert Day", "Value": peak_alert_day, "Commentary": "Dzien z najwyzsza liczba alertow."},
            {"Metric": "Reference Week", "Value": reference.get("reference_week_label", "n/a"), "Commentary": "Biezacy tydzien referencyjny."},
            {"Metric": "Reference Week Delta", "Value": reference.get("reference_release_delta", "+0"), "Commentary": "Zmiana release dla tygodnia referencyjnego."},
            {"Metric": "Weekly Alert Count", "Value": f"{weekly_alerts:,}", "Commentary": "Tygodnie oznaczone alertem."},
        ]
    )


def _classify_weekly_risk(row):
    if bool(row.get("Any Weekly Alert", False)):
        return "Critical"
    if str(row.get("Week Status", "")).lower() in {"partial range", "open week"}:
        return "Watch"
    if abs(float(row.get("Delta", 0.0))) > 0:
        return "Change"
    return "Stable"


def _build_weekly_risk_report(weekly_summary):
    if weekly_summary.empty:
        return pd.DataFrame()

    report = weekly_summary[
        [
            "Week Label",
            "Week Start",
            "Week End",
            "Week Status",
            "Working_Days_PL",
            "Quantity_Prev",
            "Quantity_Curr",
            "Delta",
            "Release Percent Label",
            "WoW Delta",
            "WoW Percent Label",
            "Any Weekly Alert",
        ]
    ].copy()
    report["Risk Level"] = report.apply(_classify_weekly_risk, axis=1)
    report = report.sort_values(
        ["Any Weekly Alert", "Week Start"],
        ascending=[False, False],
    ).reset_index(drop=True)
    return report


def _build_product_change_report(product_summary):
    if product_summary.empty:
        return pd.DataFrame()

    report = product_summary.copy()
    report["Abs Delta"] = report["Delta"].abs()
    report = report[
        [
            "Part Number",
            "Part Description",
            "Quantity_Prev",
            "Quantity_Curr",
            "Delta",
            "Abs Delta",
            "Alert_Count",
            "Change Direction",
        ]
    ].sort_values(["Abs Delta", "Alert_Count"], ascending=[False, False])
    return report.reset_index(drop=True)


def _build_alert_report(filtered_df, date_basis):
    if filtered_df.empty or "Alert" not in filtered_df.columns:
        return pd.DataFrame()

    report = filtered_df[filtered_df["Alert"]].copy()
    if report.empty:
        return pd.DataFrame()

    columns = [
        "Part Number",
        "Part Description",
        date_basis,
        "Quantity_Prev",
        "Quantity_Curr",
        "Delta",
        "Percent Change",
        "Demand Status",
        "Change Direction",
    ]
    columns = [column for column in columns if column in report.columns]
    report = report[columns].sort_values([date_basis, "Delta"], ascending=[True, False])
    report = report.rename(columns={date_basis: "Analysis Date"})
    return report.reset_index(drop=True)


def _build_demand_change_report(filtered_df):
    if filtered_df.empty or "Demand Status" not in filtered_df.columns:
        return pd.DataFrame()

    normalized = filtered_df.copy()
    normalized["Demand Status"] = (
        normalized["Demand Status"].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    )
    report = (
        normalized.groupby("Demand Status", as_index=False)
        .agg(
            Rows=("Demand Status", "size"),
            Quantity_Prev=("Quantity_Prev", "sum"),
            Quantity_Curr=("Quantity_Curr", "sum"),
            Delta=("Delta", "sum"),
            Alert_Count=("Alert", "sum"),
        )
        .sort_values(["Rows", "Delta"], ascending=[False, False])
        .reset_index(drop=True)
    )
    return report


def _classify_part_risk(alert_count, abs_delta):
    if alert_count >= 3:
        return "Critical"
    if alert_count >= 1:
        return "Alert"
    if abs_delta > 0:
        return "Monitor"
    return "Stable"


def _build_top_risk_parts(filtered_df, date_basis):
    if filtered_df.empty:
        return pd.DataFrame()

    report = (
        filtered_df.groupby(["Part Number", "Part Description"], as_index=False)
        .agg(
            Quantity_Prev=("Quantity_Prev", "sum"),
            Quantity_Curr=("Quantity_Curr", "sum"),
            Delta=("Delta", "sum"),
            Alert_Count=("Alert", "sum"),
            First_Date=(date_basis, "min"),
            Last_Date=(date_basis, "max"),
        )
    )
    report["Abs Delta"] = report["Delta"].abs()
    report["Risk Level"] = report.apply(
        lambda row: _classify_part_risk(int(row["Alert_Count"]), float(row["Abs Delta"])),
        axis=1,
    )
    report["Risk Score"] = report["Alert_Count"] * 1000000 + report["Abs Delta"]
    report = report.sort_values(["Risk Score", "Delta"], ascending=[False, False])
    return report.reset_index(drop=True)


def _build_top_risk_chart(top_risk_report):
    if top_risk_report.empty:
        return None

    chart_source = top_risk_report.head(12).copy()
    chart_source["Label"] = chart_source["Part Number"] + " | " + chart_source["Part Description"].str.slice(0, 28)
    chart_source["Color"] = chart_source["Risk Level"].map(
        {
            "Critical": "#ff6b6b",
            "Alert": "#ff9f43",
            "Monitor": "#60a5fa",
            "Stable": "#34d399",
        }
    ).fillna("#94a3b8")

    return (
        alt.Chart(chart_source)
        .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6)
        .encode(
            x=alt.X("Risk Score:Q", title="Risk Score"),
            y=alt.Y("Label:N", sort="-x", title=None),
            color=alt.Color("Color:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("Part Number:N"),
                alt.Tooltip("Part Description:N"),
                alt.Tooltip("Alert_Count:Q", title="Alert Count"),
                alt.Tooltip("Delta:Q", format=",.0f"),
                alt.Tooltip("Risk Level:N"),
            ],
        )
        .properties(height=340)
    )


def _build_demand_change_chart(demand_report):
    if demand_report.empty:
        return None

    return (
        alt.Chart(demand_report)
        .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, opacity=0.9)
        .encode(
            x=alt.X("Demand Status:N", title=None, sort="-y"),
            y=alt.Y("Rows:Q", title="Rows"),
            color=alt.value("#60a5fa"),
            tooltip=[
                alt.Tooltip("Demand Status:N"),
                alt.Tooltip("Rows:Q"),
                alt.Tooltip("Quantity_Curr:Q", format=",.0f"),
                alt.Tooltip("Delta:Q", format=",.0f"),
            ],
        )
        .properties(height=320)
    )


def _build_report_tables(data):
    return {
        "Executive Report": _build_executive_report(
            data.filtered_df,
            data.product_summary,
            data.date_summary,
            data.weekly_summary,
            data.reference,
        ),
        "Weekly Risk Report": _build_weekly_risk_report(data.weekly_summary),
        "Product Change Report": _build_product_change_report(data.product_summary),
        "Alert Report": _build_alert_report(data.filtered_df, data.date_basis),
        "Demand Change Report": _build_demand_change_report(data.filtered_df),
        "Top Risk Parts": _build_top_risk_parts(data.filtered_df, data.date_basis),
    }


def _render_report_exports(selected_label, selected_report, report_tables):
    selected_export = _format_report_table(selected_report)
    selected_csv = selected_export.to_csv(index=False).encode("utf-8") if not selected_export.empty else b""
    workbook_tables = {
        label: _format_report_table(table)
        for label, table in report_tables.items()
        if table is not None and not table.empty
    }
    workbook_bytes = _make_excel_bytes(workbook_tables) if workbook_tables else b""

    download_left, download_right = st.columns(2)
    with download_left:
        st.download_button(
            f"Pobierz {selected_label} CSV",
            data=selected_csv,
            file_name=f"{selected_label.lower().replace(' ', '_')}.csv",
            mime="text/csv",
            disabled=selected_export.empty,
            use_container_width=True,
        )
    with download_right:
        st.download_button(
            "Pobierz Reports Pack Excel",
            data=workbook_bytes,
            file_name="reports_pack.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=not workbook_tables,
            use_container_width=True,
        )


def render(data, ui):
    filtered_df = data.filtered_df
    if filtered_df.empty:
        st.info("Brak danych dla modulu Reports w aktywnych filtrach.")
        return

    report_tables = _build_report_tables(data)
    available_views = list(REPORT_VIEWS.keys())
    if report_tables["Demand Change Report"].empty:
        available_views = [view for view in available_views if view != "demand_change"]

    selected_view = st.segmented_control(
        "Raport",
        options=available_views,
        selection_mode="single",
        default=available_views[0],
        required=True,
        key="reports_module_view",
        format_func=lambda value: REPORT_VIEWS.get(value, value),
        width="stretch",
    )
    selected_view = selected_view or available_views[0]

    if selected_view == "executive":
        ui.render_section_header(
            "Reports",
            "Executive Report",
            "Zbiorczy raport zarzadzczy pokazujacy skale zmian, alerty i kluczowe punkty analizy w jednym miejscu.",
        )
        executive_report = report_tables["Executive Report"]
        ui.render_chart_table_switch(
            "reports_executive_trend",
            _interactive_chart(ui.build_quantity_chart(data.date_summary, ui.get_date_label(data.date_basis))),
            data.date_summary,
            table_height=320,
        )
        st.dataframe(_format_report_table(executive_report), use_container_width=True, height=420)
        _render_report_exports("Executive Report", executive_report, report_tables)
        return

    if selected_view == "weekly_risk":
        ui.render_section_header(
            "Reports",
            "Weekly Risk Report",
            "Raport tygodniowy skupiony na tygodniach z alertami, otwartych zakresach i istotnych zmianach release-over-release.",
        )
        weekly_report = report_tables["Weekly Risk Report"]
        ui.render_chart_table_switch(
            "reports_weekly_risk",
            _interactive_chart(ui.build_weekly_delta_chart(data.weekly_summary)),
            _format_report_table(weekly_report),
            chart_empty_message="Brak danych tygodniowych do raportu ryzyka.",
            table_height=360,
        )
        st.dataframe(_format_report_table(weekly_report), use_container_width=True, height=420)
        _render_report_exports("Weekly Risk Report", weekly_report, report_tables)
        return

    if selected_view == "product_change":
        ui.render_section_header(
            "Reports",
            "Product Change Report",
            "Raport produktowy porzadkujacy najwieksze wzrosty i spadki po materialach.",
        )
        product_report = report_tables["Product Change Report"]
        increase_chart, _ = ui.build_product_bar_chart(data.product_summary, "increase")
        ui.render_chart_table_switch(
            "reports_product_change",
            _interactive_chart(increase_chart),
            _format_report_table(product_report.head(30)),
            chart_empty_message="Brak danych produktowych do raportu zmian.",
            table_height=360,
        )
        st.dataframe(_format_report_table(product_report), use_container_width=True, height=420)
        _render_report_exports("Product Change Report", product_report, report_tables)
        return

    if selected_view == "alert":
        ui.render_section_header(
            "Reports",
            "Alert Report",
            "Raport pozycji przekraczajacych prog alertu, gotowy do szybkiego przekazania do dalszej walidacji.",
        )
        alert_report = report_tables["Alert Report"]
        ui.render_alerts(ui.build_alert_items(data.filtered_df, data.key_findings))
        st.dataframe(_format_report_table(alert_report), use_container_width=True, height=460)
        _render_report_exports("Alert Report", alert_report, report_tables)
        return

    if selected_view == "demand_change":
        ui.render_section_header(
            "Reports",
            "Demand Change Report",
            "Raport zmian popytu na bazie pola Demand Status, jesli jest dostepne w aktywnym zakresie danych.",
        )
        demand_report = report_tables["Demand Change Report"]
        ui.render_chart_table_switch(
            "reports_demand_change",
            _interactive_chart(ui.apply_chart_theme(_build_demand_change_chart(demand_report))) if not demand_report.empty else None,
            _format_report_table(demand_report),
            chart_empty_message="Brak danych Demand Status do raportu.",
            table_height=320,
        )
        st.dataframe(_format_report_table(demand_report), use_container_width=True, height=380)
        _render_report_exports("Demand Change Report", demand_report, report_tables)
        return

    ui.render_section_header(
        "Reports",
        "Top Risk Parts",
        "Raport priorytetowych materialow laczacy alerty oraz skale zmiany wolumenu.",
    )
    top_risk_report = report_tables["Top Risk Parts"]
    ui.render_chart_table_switch(
        "reports_top_risk",
        _interactive_chart(ui.apply_chart_theme(_build_top_risk_chart(top_risk_report))) if not top_risk_report.empty else None,
        _format_report_table(top_risk_report.head(25)),
        chart_empty_message="Brak danych do rankingu ryzyka.",
        table_height=360,
    )
    st.dataframe(_format_report_table(top_risk_report), use_container_width=True, height=420)
    _render_report_exports("Top Risk Parts", top_risk_report, report_tables)
