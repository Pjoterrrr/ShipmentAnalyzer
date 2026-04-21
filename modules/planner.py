import pandas as pd
import streamlit as st

from planner_helpers import (
    build_planner_coverage_chart,
    build_planner_daily_display,
    build_planner_display_table,
    build_planner_excel_bytes,
    build_planner_input_frame,
    build_planner_kpis,
    build_planner_priority_chart,
    calculate_planner_outputs,
    planner_inputs_to_state,
)


def _get_planner_storage_key(curr_meta):
    file_name = str(curr_meta.get("file_name", "planner")).strip().lower()
    return f"planner_inputs::{file_name}"


def render(data, ui):
    planner_source = data.planner_source
    curr_meta = data.curr_meta
    is_read_only = data.module_access == "read"

    ui.render_section_header(
        "Planner",
        "Planowanie produkcji wzgledem Ship Date",
        "Part Number i Part Description sa pobierane automatycznie z release'u. Operator wpisuje tylko Stock oraz opcjonalny Safety Stock.",
    )

    if planner_source.empty:
        st.info(
            "Brak dodatniego demandu w aktualnym zakresie Ship Date. Poszerz zakres dat albo wybor produktow, aby uruchomic Planner."
        )
        return

    storage_key = _get_planner_storage_key(curr_meta)
    stored_inputs = st.session_state.get(storage_key, {})
    planner_input_df = build_planner_input_frame(planner_source, stored_inputs)
    editor_key = f"{storage_key}::editor"

    planner_caption = (
        "Planner liczy wylacznie na podstawie Ship Date oraz Quantity_Curr. Filtry zakresu dat, produktow i wyszukiwarka pozostaja aktywne."
    )
    if is_read_only:
        planner_caption += " Tryb tej roli jest read-only, wiec pola Stock i Safety Stock sa zablokowane do edycji."
    st.caption(planner_caption)
    edited_inputs = st.data_editor(
        planner_input_df,
        key=editor_key,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        disabled=True if is_read_only else ["Part Number", "Part Description"],
        column_config={
            "Part Number": st.column_config.TextColumn("Part Number", width="medium"),
            "Part Description": st.column_config.TextColumn("Part Description", width="large"),
            "Stock": st.column_config.NumberColumn("Stock", min_value=0.0, step=1.0, format="%.0f"),
            "Safety Stock": st.column_config.NumberColumn("Safety Stock", min_value=0.0, step=1.0, format="%.0f"),
        },
    )
    edited_inputs["Stock"] = pd.to_numeric(edited_inputs["Stock"], errors="coerce").fillna(0.0)
    edited_inputs["Safety Stock"] = pd.to_numeric(edited_inputs["Safety Stock"], errors="coerce").fillna(0.0)
    if not is_read_only:
        st.session_state[storage_key] = planner_inputs_to_state(edited_inputs)

    planner_results, planner_daily = calculate_planner_outputs(planner_source, edited_inputs)
    planner_results_table = build_planner_display_table(planner_results)

    planner_kpis = build_planner_kpis(planner_results)
    planner_priority_chart = build_planner_priority_chart(planner_results)
    planner_coverage_chart = build_planner_coverage_chart(planner_results)
    ui.render_kpi_cards(
        [
            {
                "label": "Produkty w plannerze",
                "value": f"{planner_kpis['products']:,}",
                "copy": "Materialy z dodatnim demandem w aktualnym zakresie Ship Date.",
                "tone": "neutral",
            },
            {
                "label": "Pozycje krytyczne",
                "value": f"{planner_kpis['critical']:,}",
                "copy": "Status Krytyczne lub Wysokie ryzyko.",
                "tone": "negative",
            },
            {
                "label": "Qty To Produce Now",
                "value": f"{planner_kpis['to_produce']:,.0f}",
                "copy": "Laczna ilosc brakujaca do zabezpieczenia popytu i safety stock.",
                "tone": "positive" if planner_kpis["to_produce"] <= 0 else "negative",
            },
            {
                "label": "Sredni Coverage %",
                "value": f"{planner_kpis['avg_coverage']:.1f}%",
                "copy": f"Pokryte produkty: {planner_kpis['covered_share']:.1f}%",
                "tone": "neutral",
            },
        ]
    )

    planner_chart_left, planner_chart_right = st.columns(2, gap="large")
    with planner_chart_left:
        ui.render_chart_table_switch(
            "planner_priority",
            ui.apply_chart_theme(planner_priority_chart) if planner_priority_chart is not None else None,
            planner_results_table,
            chart_empty_message="Brak danych do rankingu Planner.",
            table_height=360,
        )
    with planner_chart_right:
        ui.render_chart_table_switch(
            "planner_coverage",
            ui.apply_chart_theme(planner_coverage_chart) if planner_coverage_chart is not None else None,
            planner_results_table,
            chart_empty_message="Brak danych do wykresu coverage.",
            table_height=360,
        )

    st.subheader("Wyniki Planner")
    st.dataframe(planner_results_table, use_container_width=True, height=420)

    selected_planner_part = st.selectbox(
        "Szczegol produktu dzien po dniu",
        options=planner_results["Part Number"].tolist(),
        format_func=lambda value: (
            f"{value} | {planner_results.loc[planner_results['Part Number'] == value, 'Part Description'].iloc[0]}"
        ),
    )
    planner_daily_detail = build_planner_daily_display(planner_daily, selected_planner_part)
    st.dataframe(planner_daily_detail, use_container_width=True, height=360)

    planner_csv_bytes = planner_results_table.to_csv(index=False).encode("utf-8")
    planner_excel_bytes = build_planner_excel_bytes(edited_inputs, planner_results, planner_daily)
    download_left, download_right = st.columns(2)
    with download_left:
        st.download_button(
            "Pobierz Planner CSV",
            data=planner_csv_bytes,
            file_name="planner_summary.csv",
            mime="text/csv",
        )
    with download_right:
        st.download_button(
            "Pobierz Planner Excel",
            data=planner_excel_bytes,
            file_name="planner_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
