import streamlit as st


REPORT_VIEWS = {
    "weekly": "Weekly",
    "product": "Product",
    "matrix": "Matrix",
}


def render(data, ui):
    filtered_df = data["filtered_df"]
    product_summary = data["product_summary"]
    weekly_summary = data["weekly_summary"]
    date_basis = data["date_basis"]
    selected_start_date = data["selected_start_date"]
    selected_end_date = data["selected_end_date"]
    reference = data["reference"]

    if filtered_df.empty:
        st.info("Brak danych dla modułu Reports w aktywnych filtrach.")
        return

    report_view = st.segmented_control(
        "Widok raportu",
        options=list(REPORT_VIEWS.keys()),
        selection_mode="single",
        default="weekly",
        required=True,
        key="reports_module_view",
        format_func=lambda value: REPORT_VIEWS.get(value, value),
        width="stretch",
    )
    report_view = report_view or "weekly"

    if report_view == "weekly":
        ui.render_section_header(
            "Weekly View",
            "Analiza tygodniowa oparta na datach",
            "Ten widok agreguje dzienne dane do poziomu tygodni ISO i ułatwia porównanie release-over-release oraz week-over-week.",
        )
        weekly_partial = weekly_summary[
            weekly_summary["Is Partial Range"] | ~weekly_summary["Is Closed Week"]
        ]
        if not weekly_partial.empty:
            st.info(
                "W tabeli i wykresach tygodnie oznaczone jako 'Partial range' lub 'Open week' obejmują niepełny zakres albo nie były jeszcze zakończone względem daty referencyjnej."
            )

        ui.render_chart_table_switch(
            "weekly_quantity",
            ui.build_weekly_quantity_chart(weekly_summary),
            ui.prepare_weekly_display_table(weekly_summary),
            chart_empty_message="Brak danych tygodniowych do wykresu.",
            table_height=360,
        )

        weekly_left, weekly_right = st.columns([1.3, 1], gap="large")
        with weekly_left:
            ui.render_chart_table_switch(
                "weekly_delta",
                ui.build_weekly_delta_chart(weekly_summary),
                ui.prepare_weekly_display_table(weekly_summary),
                chart_empty_message="Brak danych tygodniowych do wykresu delta.",
                table_height=320,
            )
        with weekly_right:
            weekly_focus = ui.build_weekly_focus_table(
                weekly_summary,
                reference["reference_week_label"],
                reference["previous_week_label"],
                reference["reference_release_delta"],
                reference["reference_release_pct"],
                reference["reference_wow_delta"],
                reference["reference_wow_pct"],
            )
            st.subheader("Porównanie tygodni")
            st.dataframe(weekly_focus, use_container_width=True, height=240)

        st.subheader("Tabela tygodniowa")
        st.dataframe(ui.prepare_weekly_display_table(weekly_summary), use_container_width=True, height=420)
        return

    if report_view == "product":
        if product_summary.empty:
            st.info("Brak danych produktowych dla aktywnych filtrów.")
            return

        ui.render_section_header(
            "Product Drilldown",
            "Analiza wybranego produktu",
            "Skup się na jednym materiale i prześledź jego ruch po dniach oraz tygodniach bez utraty kontekstu filtrowania.",
        )
        selected_product_label = st.selectbox(
            "Wybierz produkt",
            options=product_summary["Product Label"].tolist(),
        )
        product_detail = filtered_df[
            filtered_df["Product Label"] == selected_product_label
        ].sort_values(date_basis)
        product_date_summary = ui.summarize_dates(product_detail, date_basis)

        product_metrics = st.columns(4)
        product_metrics[0].metric(
            "Poprzednia ilość", f"{product_detail['Quantity_Prev'].sum():,.0f}"
        )
        product_metrics[1].metric(
            "Aktualna ilość", f"{product_detail['Quantity_Curr'].sum():,.0f}"
        )
        product_metrics[2].metric(
            "Bilans zmian", f"{product_detail['Delta'].sum():+,.0f}"
        )
        product_metrics[3].metric("Liczba alertów", int(product_detail["Alert"].sum()))

        ui.render_chart_table_switch(
            "product_quantity",
            ui.build_quantity_chart(product_date_summary, ui.get_date_label(date_basis)),
            product_date_summary,
            table_height=320,
        )
        ui.render_chart_table_switch(
            "product_delta",
            ui.build_delta_chart(product_date_summary, ui.get_date_label(date_basis)),
            product_date_summary,
            table_height=320,
        )

        product_weekly_summary = ui.build_weekly_summary(
            product_detail,
            date_basis,
            selected_start_date,
            selected_end_date,
            selected_end_date,
            ui.threshold,
        )
        st.subheader("Tygodnie ISO dla produktu")
        ui.render_chart_table_switch(
            "product_weekly",
            ui.build_weekly_quantity_chart(product_weekly_summary),
            ui.prepare_weekly_display_table(product_weekly_summary),
            chart_empty_message="Brak danych tygodniowych dla wybranego produktu.",
            table_height=280,
        )

        product_table = ui.build_product_detail_table(product_detail)
        st.dataframe(product_table, use_container_width=True, height=360)
        return

    ui.render_section_header(
        "Release Matrix",
        "Macierz podobna do arkusza release'u",
        "Macierz zachowuje układ bliski pracy w Excelu, ale pozostaje spójna wizualnie z całym dashboardem.",
    )
    matrix_metric = st.segmented_control(
        "Metryka",
        options=["Current Quantity", "Previous Quantity", "Delta", "Percent Change"],
        selection_mode="single",
        default="Current Quantity",
        required=True,
        key="reports_matrix_metric",
        format_func=ui.get_metric_label,
        width="stretch",
    )
    matrix_metric = matrix_metric or "Current Quantity"
    matrix = ui.build_matrix(filtered_df, date_basis, matrix_metric)
    matrix_cells = matrix.shape[0] * max(matrix.shape[1], 1)

    if matrix.empty:
        st.info("Brak danych do macierzy.")
    elif matrix_cells <= ui.max_matrix_style_cells:
        st.dataframe(
            ui.style_matrix(matrix, matrix_metric),
            use_container_width=True,
            height=520,
        )
    else:
        st.info(
            "Macierz jest zbyt duża do stylowania, dlatego pokazuję ją bez dodatkowego formatowania."
        )
        st.dataframe(matrix, use_container_width=True, height=520)
