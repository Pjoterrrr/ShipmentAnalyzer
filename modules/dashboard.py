import streamlit as st


def render(data, ui):
    filtered_df = data.filtered_df
    product_summary = data.product_summary
    date_summary = data.date_summary
    weekly_summary = data.weekly_summary
    key_findings = data.key_findings
    date_basis = data.date_basis
    reference = data.reference

    if filtered_df.empty:
        st.info("Brak danych dla modulu Dashboard w aktywnych filtrach.")
        return

    ui.render_section_header(
        "Dashboard",
        "Executive overview",
        "Najwazniejsze KPI, trendy i produkty wymagajace uwagi w jednym enterprise widoku.",
    )
    ui.render_kpi_row(ui.build_dashboard_kpi_metrics(filtered_df, product_summary, date_summary))

    ui.render_section_header(
        "Alerts & Insights",
        "Priorytety do sprawdzenia",
        "Najwazniejsze sygnaly, ktore warto zweryfikowac w pierwszej kolejnosci.",
    )
    ui.render_alerts(ui.build_alert_items(filtered_df, key_findings))

    ui.render_section_header(
        "Reference Week",
        "Szybki odczyt tygodniowy",
        (
            f"Analiza tygodniowa odnosi sie do {reference['reference_week_label']} "
            f"({reference['reference_range_label']}). Data referencyjna: {data.selected_end_date:%Y-%m-%d}."
        ),
    )
    ui.render_kpi_row(
        [
            {
                "label": "Wolumen tygodnia",
                "value": reference["reference_curr_qty"],
                "delta_label": "Release",
                "delta": reference["reference_release_delta"],
                "copy": "Biezacy wolumen tygodnia referencyjnego.",
                "accent": "#2d81ff",
                "delta_width": 88,
            },
            {
                "label": "Zmiana vs release",
                "value": reference["reference_release_pct"],
                "delta_label": "Baseline",
                "delta": reference["reference_prev_qty"],
                "copy": "Porownanie do poprzedniego release'u.",
                "accent": "#00c4b4",
                "delta_width": 64,
            },
            {
                "label": "Zmiana WoW",
                "value": reference["reference_wow_delta"],
                "delta_label": "WoW %",
                "delta": reference["reference_wow_pct"],
                "copy": f"Wzgledem {reference['previous_week_label']}.",
                "accent": "#8957e5",
                "delta_width": 56,
            },
            {
                "label": "Dni robocze PL",
                "value": str(reference["reference_working_days"]),
                "delta_label": "Na dzien",
                "delta": reference["reference_per_day"],
                "copy": "Kontekst produktywnosci tygodnia.",
                "accent": "#d29922",
                "delta_width": 40,
            },
        ]
    )

    ui.render_section_header(
        "Trend",
        f"Release trend wedlug osi: {ui.get_date_label(date_basis)}",
        "Porownanie poprzedniego i aktualnego release'u z interaktywnym drill-down do danych.",
    )
    ui.render_chart_table_switch(
        "dashboard_trend",
        ui.build_quantity_chart(date_summary, ui.get_date_label(date_basis)),
        date_summary,
        table_height=360,
    )

    trend_left, trend_right = st.columns([1.15, 0.85], gap="large")
    with trend_left:
        ui.render_section_header(
            "Delta",
            "Bilans zmian w czasie",
            "Zielone i czerwone slupki od razu pokazują kierunek zmian w kolejnych dniach.",
        )
        ui.render_chart_table_switch(
            "dashboard_delta",
            ui.build_delta_chart(date_summary, ui.get_date_label(date_basis)),
            date_summary,
            table_height=320,
        )
    with trend_right:
        ui.render_section_header(
            "Mix",
            "Struktura zmian",
            "Szybki podzial na wzrosty, spadki i brak zmian.",
        )
        ui.render_chart_table_switch(
            "dashboard_mix",
            ui.build_change_mix_chart(filtered_df),
            ui.build_change_mix_source(filtered_df),
            table_height=240,
        )

    waterfall_left, waterfall_right = st.columns([1.25, 0.75], gap="large")
    with waterfall_left:
        ui.render_section_header(
            "Products",
            "Waterfall top zmian",
            "Najwieksze ruchy po materialach w formie syntetycznego waterfall chart.",
        )
        waterfall_chart = ui.build_product_waterfall_chart(product_summary)
        if waterfall_chart is None:
            st.info("Brak danych produktowych do waterfall chart.")
        else:
            waterfall_source = (
                product_summary.assign(Abs_Delta=product_summary["Delta"].abs())
                .sort_values("Abs_Delta", ascending=False)
                .drop(columns=["Abs_Delta"])
                .head(8)
            )
            ui.render_chart_table_switch(
                "dashboard_waterfall",
                waterfall_chart,
                waterfall_source,
                table_height=320,
            )
    with waterfall_right:
        ui.render_section_header(
            "Weekly",
            "Tygodnie ISO",
            "Konsolidacja tygodniowa dla szybkiej oceny rytmu release'ow.",
        )
        weekly_chart = ui.build_weekly_quantity_chart(weekly_summary)
        weekly_preview = ui.prepare_weekly_display_table(weekly_summary).tail(8)
        ui.render_chart_table_switch(
            "dashboard_weekly",
            weekly_chart,
            weekly_preview,
            chart_empty_message="Brak danych tygodniowych do wykresu.",
            table_height=320,
        )

    increase_chart, increase_title = ui.build_product_bar_chart(product_summary, "increase")
    decrease_chart, decrease_title = ui.build_product_bar_chart(product_summary, "decrease")
    dashboard_left, dashboard_right = st.columns(2, gap="large")

    with dashboard_left:
        ui.render_section_header("Growth", increase_title, "Produkty z najsilniejszym wzrostem wolumenu.")
        if increase_chart is None:
            st.info("Brak produktow ze wzrostem w aktualnym filtrowaniu.")
        else:
            ui.render_chart_table_switch(
                "dashboard_increase",
                increase_chart,
                ui.build_product_bar_source(product_summary, "increase"),
                table_height=340,
            )

    with dashboard_right:
        ui.render_section_header("Decline", decrease_title, "Produkty z najwiekszym spadkiem wolumenu.")
        if decrease_chart is None:
            st.info("Brak produktow ze spadkiem w aktualnym filtrowaniu.")
        else:
            ui.render_chart_table_switch(
                "dashboard_decrease",
                decrease_chart,
                ui.build_product_bar_source(product_summary, "decrease"),
                table_height=340,
            )

    ui.render_section_header(
        "Highlights",
        "Najwazniejsze zmiany",
        "Tabela dla produktow z najwyzszym bezwzglednym ruchem oraz alertami.",
    )
    highlight_table = (
        product_summary.assign(Abs_Delta=product_summary["Delta"].abs())
        .sort_values("Abs_Delta", ascending=False)
        .drop(columns=["Abs_Delta"])
        .head(10)
    )
    highlight_table["Quantity_Prev"] = highlight_table["Quantity_Prev"].map(lambda value: f"{value:,.0f}")
    highlight_table["Quantity_Curr"] = highlight_table["Quantity_Curr"].map(lambda value: f"{value:,.0f}")
    highlight_table["Delta"] = highlight_table["Delta"].map(ui.format_signed_int)
    highlight_table = highlight_table.rename(
        columns={
            "Part Number": "Numer czesci",
            "Part Description": "Opis produktu",
            "Quantity_Prev": "Poprzednia ilosc",
            "Quantity_Curr": "Aktualna ilosc",
            "Delta": "Zmiana ilosci",
            "Alert_Count": "Liczba alertow",
            "Change Direction": "Kierunek zmiany",
        }
    )
    st.dataframe(highlight_table, use_container_width=True, height=360)
