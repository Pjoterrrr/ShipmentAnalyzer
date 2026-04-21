import streamlit as st


def render(data, ui):
    filtered_df = data["filtered_df"]
    product_summary = data["product_summary"]
    date_summary = data["date_summary"]
    weekly_summary = data["weekly_summary"]
    key_findings = data["key_findings"]
    date_basis = data["date_basis"]
    reference = data["reference"]

    if filtered_df.empty:
        st.info("Brak danych dla modułu Dashboard w aktywnych filtrach.")
        return

    ui.render_section_header(
        "KPI",
        "Najważniejsze wskaźniki",
        "Karty poniżej pokazują główne liczby do szybkiego odczytu bez przeskakiwania między modułami.",
    )
    ui.render_kpi_cards(ui.build_kpi_metrics(filtered_df, product_summary))

    ui.render_section_header(
        "Alerts & Insights",
        "Priorytety do sprawdzenia",
        "Najważniejsze sygnały, które warto zweryfikować w pierwszej kolejności.",
    )
    ui.render_alerts(ui.build_alert_items(filtered_df, key_findings))

    ui.render_section_header(
        "Reference Week",
        "Szybki odczyt tygodniowy",
        (
            f"Analiza tygodniowa odnosi się do {reference['reference_week_label']} "
            f"({reference['reference_range_label']}). Data referencyjna: {data['selected_end_date']:%Y-%m-%d}."
        ),
    )
    ui.render_kpi_cards(
        [
            {
                "label": "Wolumen tygodnia",
                "value": reference["reference_curr_qty"],
                "copy": f"Bilans release: {reference['reference_release_delta']}",
                "tone": "neutral",
            },
            {
                "label": "Zmiana vs poprzedni release",
                "value": reference["reference_release_pct"],
                "copy": f"Poprzedni wolumen: {reference['reference_prev_qty']}",
                "tone": "neutral",
            },
            {
                "label": "Zmiana WoW",
                "value": reference["reference_wow_delta"],
                "copy": f"{reference['reference_wow_pct']} względem {reference['previous_week_label']}",
                "tone": "neutral",
            },
            {
                "label": "Dni robocze PL",
                "value": str(reference["reference_working_days"]),
                "copy": reference["reference_per_day"],
                "tone": "neutral",
            },
        ]
    )

    ui.render_section_header(
        "Dashboard",
        f"Trend zmian według osi: {ui.get_date_label(date_basis)}",
        "Widok główny zbiera najważniejsze wykresy, strukturę zmian oraz szybki podgląd produktów z największym ruchem.",
    )
    ui.render_chart_table_switch(
        "dashboard_trend",
        ui.build_quantity_chart(date_summary, ui.get_date_label(date_basis)),
        date_summary,
        table_height=360,
    )

    trend_left, trend_right = st.columns([1.45, 1], gap="large")
    with trend_left:
        ui.render_chart_table_switch(
            "dashboard_delta",
            ui.build_delta_chart(date_summary, ui.get_date_label(date_basis)),
            date_summary,
            table_height=320,
        )
    with trend_right:
        st.subheader("Struktura zmian")
        ui.render_chart_table_switch(
            "dashboard_mix",
            ui.build_change_mix_chart(filtered_df),
            ui.build_change_mix_source(filtered_df),
            table_height=240,
        )

    increase_chart, increase_title = ui.build_product_bar_chart(product_summary, "increase")
    decrease_chart, decrease_title = ui.build_product_bar_chart(product_summary, "decrease")
    dashboard_left, dashboard_right = st.columns(2)

    with dashboard_left:
        st.subheader(increase_title)
        if increase_chart is None:
            st.info("Brak produktów ze wzrostem w aktualnym filtrowaniu.")
        else:
            ui.render_chart_table_switch(
                "dashboard_increase",
                increase_chart,
                ui.build_product_bar_source(product_summary, "increase"),
                table_height=340,
            )

    with dashboard_right:
        st.subheader(decrease_title)
        if decrease_chart is None:
            st.info("Brak produktów ze spadkiem w aktualnym filtrowaniu.")
        else:
            ui.render_chart_table_switch(
                "dashboard_decrease",
                decrease_chart,
                ui.build_product_bar_source(product_summary, "decrease"),
                table_height=340,
            )

    st.subheader("Najważniejsze zmiany")
    highlight_table = (
        product_summary.assign(Abs_Delta=product_summary["Delta"].abs())
        .sort_values("Abs_Delta", ascending=False)
        .drop(columns=["Abs_Delta"])
        .head(10)
    )
    highlight_table["Quantity_Prev"] = highlight_table["Quantity_Prev"].map(
        lambda value: f"{value:,.0f}"
    )
    highlight_table["Quantity_Curr"] = highlight_table["Quantity_Curr"].map(
        lambda value: f"{value:,.0f}"
    )
    highlight_table["Delta"] = highlight_table["Delta"].map(ui.format_signed_int)
    highlight_table = highlight_table.rename(
        columns={
            "Part Number": "Numer części",
            "Part Description": "Opis produktu",
            "Quantity_Prev": "Poprzednia ilość",
            "Quantity_Curr": "Aktualna ilość",
            "Delta": "Zmiana ilości",
            "Alert_Count": "Liczba alertów",
            "Change Direction": "Kierunek zmiany",
        }
    )
    st.dataframe(highlight_table, use_container_width=True, height=360)

    st.subheader("Tygodnie ISO")
    weekly_chart = ui.build_weekly_quantity_chart(weekly_summary)
    weekly_preview = ui.prepare_weekly_display_table(weekly_summary).tail(8)
    ui.render_chart_table_switch(
        "dashboard_weekly",
        weekly_chart,
        weekly_preview,
        chart_empty_message="Brak danych tygodniowych do wykresu.",
        table_height=320,
    )
