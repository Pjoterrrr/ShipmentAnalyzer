import streamlit as st


def render(data, ui):
    filtered_df = data["filtered_df"]
    weekly_summary = data["weekly_summary"]
    product_summary = data["product_summary"]
    prev_meta = data["prev_meta"]
    curr_meta = data["curr_meta"]
    date_basis = data["date_basis"]
    selected_start_date = data["selected_start_date"]
    selected_end_date = data["selected_end_date"]
    key_findings = data["key_findings"]
    excel_bytes = data.get("excel_bytes")
    csv_bytes = data.get("csv_bytes")

    if filtered_df.empty:
        st.info("Brak danych szczegółowych dla aktywnych filtrów.")
        return

    ui.render_section_header(
        "Details",
        "Dane szczegółowe i eksport",
        "Pełny podgląd przefiltrowanych wierszy do szybkiej walidacji oraz eksportu do dalszej pracy operacyjnej.",
    )
    preview_limit = st.selectbox(
        "Liczba wierszy w podglądzie",
        options=[100, 250, 500, 1000],
        index=2,
        key="details_preview_limit",
    )
    detail_table = ui.build_detail_export_table(filtered_df)

    if len(detail_table) > preview_limit:
        st.info(
            f"Pokazuje pierwsze {preview_limit} z {len(detail_table)} wierszy. Pełny raport jest dostępny do pobrania."
        )
    st.dataframe(
        detail_table.head(preview_limit),
        use_container_width=True,
        height=420,
    )

    if excel_bytes is None:
        current_matrix_for_export = ui.build_matrix(filtered_df, date_basis, "Current Quantity")
        delta_matrix_for_export = ui.build_matrix(filtered_df, date_basis, "Delta")
        excel_bytes = ui.to_excel_bytes(
            filtered_df,
            weekly_summary,
            current_matrix_for_export,
            delta_matrix_for_export,
            prev_meta,
            curr_meta,
            product_summary,
            date_basis,
            selected_start_date,
            selected_end_date,
            key_findings,
        )
    if csv_bytes is None:
        csv_bytes = detail_table.to_csv(index=False).encode("utf-8")

    download_left, download_right = st.columns(2)
    with download_left:
        st.download_button(
            "Pobierz filtrowane dane CSV",
            data=csv_bytes,
            file_name="pjoter_development_release_change_filtered.csv",
            mime="text/csv",
        )
    with download_right:
        st.download_button(
            "Pobierz raport Excel",
            data=excel_bytes,
            file_name="pjoter_development_release_change_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
