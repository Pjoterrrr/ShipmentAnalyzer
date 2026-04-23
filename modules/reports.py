import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from analytics_calendar import build_calendar_frame


REPORT_VIEWS = {
    "weekly_quantity": "Weekly quantity comparison",
    "delta_heatmap": "Delta heat map",
    "current_matrix": "Current matrix",
    "operational_calendar": "Polish operational calendar",
}


def _format_report_table(dataframe):
    if dataframe is None or dataframe.empty:
        return pd.DataFrame()

    table = dataframe.copy()
    for column in table.columns:
        if pd.api.types.is_datetime64_any_dtype(table[column]):
            table[column] = pd.to_datetime(table[column], errors="coerce").dt.strftime("%Y-%m-%d")
    return table


def _direction_label(value):
    numeric_value = float(value or 0)
    if numeric_value > 0:
        return "Wzrost"
    if numeric_value < 0:
        return "Spadek"
    return "Stabilnie"


def _part_label(part_number, part_description):
    description = str(part_description or "").strip()
    return str(part_number) if not description else f"{part_number} | {description}"


def build_report_dataset(data):
    weekly_comparison = generate_weekly_comparison_report(data)
    delta_heatmap = generate_delta_heatmap_report(data)
    current_matrix = generate_current_matrix_report(data)
    operational_calendar = build_polish_operational_calendar(data)
    calendar_summary = _build_calendar_weekly_summary(operational_calendar)
    return {
        "weekly_comparison": weekly_comparison,
        "delta_heatmap": delta_heatmap,
        "current_matrix": current_matrix,
        "operational_calendar": operational_calendar,
        "calendar_summary": calendar_summary,
    }


def generate_weekly_comparison_report(data):
    if data.weekly_summary is None or data.weekly_summary.empty:
        return pd.DataFrame()

    report = data.weekly_summary[
        [
            "Week Label",
            "Week Start",
            "Week End",
            "Quantity_Prev",
            "Quantity_Curr",
            "Delta",
            "Release Percent Label",
            "Previous Week Current Qty",
            "WoW Delta",
            "WoW Percent Label",
            "Working_Days_PL",
            "Holidays_PL",
            "Week Status",
        ]
    ].copy()
    report["Trend"] = report["WoW Delta"].map(_direction_label)
    report["Holiday Week"] = report["Holidays_PL"].map(lambda value: "Tak" if int(value or 0) > 0 else "")
    report["Week Start"] = pd.to_datetime(report["Week Start"], errors="coerce").dt.strftime("%Y-%m-%d")
    report["Week End"] = pd.to_datetime(report["Week End"], errors="coerce").dt.strftime("%Y-%m-%d")
    report = report.rename(
        columns={
            "Quantity_Prev": "Previous Release Qty",
            "Quantity_Curr": "Current Release Qty",
            "Delta": "Release Delta",
            "Release Percent Label": "Release Change %",
            "Previous Week Current Qty": "Previous Week Qty",
            "WoW Delta": "WoW Delta",
            "WoW Percent Label": "WoW Change %",
            "Working_Days_PL": "Working Days PL",
            "Holidays_PL": "Polish Holidays",
        }
    )
    return report.reset_index(drop=True)


def _build_part_date_matrix(filtered_df, date_basis, metric_name):
    if filtered_df is None or filtered_df.empty:
        return pd.DataFrame()

    metric_column = {
        "Current Quantity": "Quantity_Curr",
        "Delta": "Delta",
    }.get(metric_name)
    if metric_column is None:
        return pd.DataFrame()

    source = filtered_df.copy()
    source[date_basis] = pd.to_datetime(source[date_basis], errors="coerce")
    source = source.dropna(subset=[date_basis])
    if source.empty:
        return pd.DataFrame()

    grouped = (
        source.groupby(["Part Number", "Part Description", date_basis], as_index=False)[metric_column]
        .sum()
    )
    matrix = grouped.pivot_table(
        index=["Part Number", "Part Description"],
        columns=date_basis,
        values=metric_column,
        aggfunc="sum",
        fill_value=0,
    )
    matrix = matrix.sort_index(axis=1)
    matrix.columns = [pd.Timestamp(column).strftime("%Y-%m-%d") for column in matrix.columns]
    matrix = matrix.reset_index()
    return matrix


def generate_delta_heatmap_report(data):
    return _build_part_date_matrix(data.filtered_df, data.date_basis, "Delta")


def generate_current_matrix_report(data):
    return _build_part_date_matrix(data.filtered_df, data.date_basis, "Current Quantity")


def build_polish_operational_calendar(data):
    calendar = build_calendar_frame(data.selected_start_date, data.selected_end_date)
    if calendar.empty:
        return pd.DataFrame()

    calendar = calendar.copy()
    calendar["Date"] = pd.to_datetime(calendar["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    calendar["Week Start"] = pd.to_datetime(calendar["Week Start"], errors="coerce").dt.strftime("%Y-%m-%d")
    calendar["Week End"] = pd.to_datetime(calendar["Week End"], errors="coerce").dt.strftime("%Y-%m-%d")
    return calendar[
        [
            "Date",
            "Day Name",
            "Day Type",
            "Holiday Name",
            "ISO Week Label",
            "Week Start",
            "Week End",
            "Is Working Day",
            "Is Weekend",
            "Is Holiday",
        ]
    ]


def _build_calendar_weekly_summary(calendar_df):
    if calendar_df is None or calendar_df.empty:
        return pd.DataFrame()

    grouped = (
        calendar_df.groupby(["ISO Week Label", "Week Start", "Week End"], as_index=False)
        .agg(
            Working_Days_PL=("Is Working Day", "sum"),
            Saturdays=("Day Type", lambda values: int((pd.Series(values) == "Saturday").sum())),
            Sundays=("Day Type", lambda values: int((pd.Series(values) == "Sunday").sum())),
            Polish_Holidays=("Is Holiday", "sum"),
        )
        .sort_values(["Week Start", "Week End"])
        .reset_index(drop=True)
    )
    grouped["Non-working Days"] = grouped["Saturdays"] + grouped["Sundays"] + grouped["Polish_Holidays"]
    return grouped


def _build_weekly_quantity_chart(weekly_report):
    if weekly_report.empty:
        return None

    figure = go.Figure()
    figure.add_trace(
        go.Scatter(
            x=weekly_report["Week Label"],
            y=weekly_report["Previous Release Qty"],
            mode="lines+markers",
            name="Previous release",
            line={"color": "#8b949e", "width": 2},
            marker={"size": 8},
            hovertemplate="Week: %{x}<br>Previous release: %{y:,.0f}<extra></extra>",
        )
    )
    figure.add_trace(
        go.Scatter(
            x=weekly_report["Week Label"],
            y=weekly_report["Current Release Qty"],
            mode="lines+markers",
            name="Current release",
            line={"color": "#2d81ff", "width": 3},
            marker={"size": 9},
            customdata=weekly_report[["WoW Delta", "WoW Change %", "Working Days PL", "Holiday Week"]].to_numpy(),
            hovertemplate=(
                "Week: %{x}<br>"
                "Current release: %{y:,.0f}<br>"
                "WoW delta: %{customdata[0]:+,.0f}<br>"
                "WoW change: %{customdata[1]}<br>"
                "Working days PL: %{customdata[2]}<br>"
                "Holiday week: %{customdata[3]}<extra></extra>"
            ),
        )
    )
    figure.add_trace(
        go.Bar(
            x=weekly_report["Week Label"],
            y=weekly_report["WoW Delta"],
            name="WoW delta",
            marker={
                "color": [
                    "#3fb950" if value > 0 else "#f85149" if value < 0 else "#8b949e"
                    for value in weekly_report["WoW Delta"]
                ]
            },
            opacity=0.28,
            hovertemplate="Week: %{x}<br>WoW delta: %{y:+,.0f}<extra></extra>",
        )
    )
    figure.update_layout(height=380, barmode="relative", xaxis_title=None, yaxis_title="Qty")
    return figure


def _build_matrix_heatmap(matrix_report, metric_name):
    if matrix_report.empty:
        return None

    value_columns = [
        column for column in matrix_report.columns if column not in {"Part Number", "Part Description"}
    ]
    if not value_columns:
        return None

    chart_source = matrix_report.copy()
    chart_source["Label"] = chart_source.apply(
        lambda row: _part_label(row["Part Number"], str(row["Part Description"])[:36]),
        axis=1,
    )
    chart_source["Magnitude"] = chart_source[value_columns].abs().sum(axis=1)
    chart_source = chart_source.sort_values("Magnitude", ascending=False).head(25)
    color_scale = "RdBu" if metric_name == "Delta" else "Blues"
    z_mid = 0 if metric_name == "Delta" else None

    figure = go.Figure(
        data=[
            go.Heatmap(
                z=chart_source[value_columns].to_numpy(),
                x=value_columns,
                y=chart_source["Label"],
                colorscale=color_scale,
                zmid=z_mid,
                colorbar={"title": metric_name},
                hovertemplate="Part: %{y}<br>Date: %{x}<br>Value: %{z:,.0f}<extra></extra>",
            )
        ]
    )
    figure.update_layout(height=680, xaxis_title=None, yaxis_title=None)
    return figure


def _build_calendar_chart(calendar_summary):
    if calendar_summary.empty:
        return None

    figure = go.Figure()
    figure.add_trace(
        go.Bar(
            x=calendar_summary["ISO Week Label"],
            y=calendar_summary["Working_Days_PL"],
            name="Working days PL",
            marker={"color": "#3fb950"},
            hovertemplate="Week: %{x}<br>Working days PL: %{y}<extra></extra>",
        )
    )
    figure.add_trace(
        go.Bar(
            x=calendar_summary["ISO Week Label"],
            y=calendar_summary["Saturdays"],
            name="Saturdays",
            marker={"color": "#d29922"},
            hovertemplate="Week: %{x}<br>Saturdays: %{y}<extra></extra>",
        )
    )
    figure.add_trace(
        go.Bar(
            x=calendar_summary["ISO Week Label"],
            y=calendar_summary["Polish_Holidays"],
            name="Polish holidays",
            marker={"color": "#f85149"},
            hovertemplate="Week: %{x}<br>Polish holidays: %{y}<extra></extra>",
        )
    )
    figure.update_layout(height=360, barmode="group", xaxis_title=None, yaxis_title="Days")
    return figure


def _render_reports_context(data):
    st.caption(
        "Raporty korzystaja wyłącznie z aktywnego zakresu filtrowania: "
        f"{data.selected_start_date:%Y-%m-%d} - {data.selected_end_date:%Y-%m-%d}. "
        f"Wiersze po filtrach: {len(data.filtered_df):,}."
    )


def render(data, ui):
    if data.filtered_df is None or data.filtered_df.empty:
        st.info("Brak danych dla modulu Reports w aktywnych filtrach.")
        return

    report_dataset = build_report_dataset(data)
    selected_view = st.segmented_control(
        "Raport",
        options=list(REPORT_VIEWS.keys()),
        selection_mode="single",
        default="weekly_quantity",
        required=True,
        key="reports_module_view",
        format_func=lambda value: REPORT_VIEWS.get(value, value),
        width="stretch",
    )
    selected_view = selected_view or "weekly_quantity"

    if selected_view == "weekly_quantity":
        ui.render_section_header(
            "Reports",
            "Weekly quantity comparison",
            "Najwazniejszy raport operacyjny: ilosci tydzien do tygodnia, zmiana procentowa oraz kontekst dni roboczych i swiat PL.",
        )
        _render_reports_context(data)
        weekly_report = report_dataset["weekly_comparison"]
        ui.render_chart_table_switch(
            "reports_weekly_quantity",
            _build_weekly_quantity_chart(weekly_report),
            weekly_report,
            chart_empty_message="Brak danych tygodniowych dla aktywnego zakresu.",
            table_height=420,
        )
        return

    if selected_view == "delta_heatmap":
        ui.render_section_header(
            "Reports",
            "Delta heat map",
            "Mapa zmian po part number dla przefiltrowanego zakresu. Pokazuje tylko indeksy pozostajace po aktywnej filtracji.",
        )
        _render_reports_context(data)
        delta_report = report_dataset["delta_heatmap"]
        ui.render_chart_table_switch(
            "reports_delta_heatmap",
            _build_matrix_heatmap(delta_report, "Release Delta"),
            delta_report,
            chart_empty_message="Brak danych do delta heat map dla aktywnego zakresu.",
            table_height=520,
        )
        return

    if selected_view == "current_matrix":
        ui.render_section_header(
            "Reports",
            "Current matrix",
            "Macierz aktualnego wolumenu po filtrach. Raport pokazuje tylko dane, ktore przeszly aktywne filtry zakresu i produktow.",
        )
        _render_reports_context(data)
        current_matrix_report = report_dataset["current_matrix"]
        ui.render_chart_table_switch(
            "reports_current_matrix",
            _build_matrix_heatmap(current_matrix_report, "Current Release Qty"),
            current_matrix_report,
            chart_empty_message="Brak danych do current matrix dla aktywnego zakresu.",
            table_height=520,
        )
        return

    ui.render_section_header(
        "Reports",
        "Polish operational calendar",
        "Kalendarz operacyjny dla analizowanego zakresu z oznaczeniem sobot, niedziel i polskich swiat ustawowych.",
    )
    _render_reports_context(data)
    ui.render_chart_table_switch(
        "reports_operational_calendar",
        _build_calendar_chart(report_dataset["calendar_summary"]),
        report_dataset["operational_calendar"],
        chart_empty_message="Brak danych kalendarzowych dla aktywnego zakresu.",
        table_height=520,
    )
