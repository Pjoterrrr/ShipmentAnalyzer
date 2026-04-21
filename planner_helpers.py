import io

import altair as alt
import pandas as pd


PLANNER_COLUMNS = [
    "Part Number",
    "Part Description",
    "Ship Date",
    "Demand Qty",
]


def _empty_planner_source():
    return pd.DataFrame(columns=PLANNER_COLUMNS)


def _normalize_number(value):
    if value in ("", None):
        return 0.0
    numeric_value = pd.to_numeric(value, errors="coerce")
    if pd.isna(numeric_value):
        return 0.0
    return float(numeric_value)


def prepare_planner_source(dataframe):
    if dataframe is None or dataframe.empty:
        return _empty_planner_source()

    source = dataframe.copy()
    required = {"Part Number", "Part Description", "Ship Date", "Quantity_Curr"}
    if not required.issubset(source.columns):
        return _empty_planner_source()

    source["Ship Date"] = pd.to_datetime(source["Ship Date"], errors="coerce")
    source["Quantity_Curr"] = pd.to_numeric(source["Quantity_Curr"], errors="coerce").fillna(0.0)
    source["Part Number"] = source["Part Number"].fillna("").astype(str).str.strip()
    source["Part Description"] = source["Part Description"].fillna("").astype(str).str.strip()
    source = source[source["Ship Date"].notna()]
    source = source[source["Part Number"].ne("")]
    source = source[source["Quantity_Curr"] > 0]

    if source.empty:
        return _empty_planner_source()

    source = (
        source.groupby(["Part Number", "Part Description", "Ship Date"], as_index=False)
        .agg(Demand_Qty=("Quantity_Curr", "sum"))
        .rename(columns={"Demand_Qty": "Demand Qty"})
        .sort_values(["Part Number", "Ship Date"])
        .reset_index(drop=True)
    )
    return source


def build_planner_input_frame(planner_source, stored_inputs=None):
    stored_inputs = stored_inputs or {}
    if planner_source is None or planner_source.empty:
        return pd.DataFrame(
            columns=["Part Number", "Part Description", "Stock", "Safety Stock"]
        )

    products = (
        planner_source[["Part Number", "Part Description"]]
        .drop_duplicates()
        .sort_values(["Part Number", "Part Description"])
        .reset_index(drop=True)
    )

    rows = []
    for record in products.to_dict("records"):
        saved = stored_inputs.get(record["Part Number"], {})
        rows.append(
            {
                "Part Number": record["Part Number"],
                "Part Description": record["Part Description"],
                "Stock": _normalize_number(saved.get("Stock", 0.0)),
                "Safety Stock": _normalize_number(saved.get("Safety Stock", 0.0)),
            }
        )
    return pd.DataFrame(rows)


def planner_inputs_to_state(planner_input_df):
    if planner_input_df is None or planner_input_df.empty:
        return {}

    state = {}
    for row in planner_input_df.to_dict("records"):
        part_number = str(row.get("Part Number", "")).strip()
        if not part_number:
            continue
        state[part_number] = {
            "Stock": _normalize_number(row.get("Stock", 0.0)),
            "Safety Stock": _normalize_number(row.get("Safety Stock", 0.0)),
        }
    return state


def _classify_status(shortage_date, shortage_qty, coverage_pct, total_demand, today):
    if total_demand <= 0:
        return "Brak popytu"
    if shortage_qty <= 0:
        return "Pokryte"
    if pd.isna(shortage_date):
        return "Monitoruj"

    days_to_shortage = int((pd.Timestamp(shortage_date).normalize() - today).days)
    if days_to_shortage <= 3:
        return "Krytyczne"
    if days_to_shortage <= 10:
        return "Wysokie ryzyko"
    if coverage_pct < 100:
        return "Ryzyko"
    return "Monitoruj"


def _status_rank(status):
    order = {
        "Krytyczne": 0,
        "Wysokie ryzyko": 1,
        "Ryzyko": 2,
        "Monitoruj": 3,
        "Pokryte": 4,
        "Brak popytu": 5,
    }
    return order.get(status, 9)


def _build_recommendation(status, shortage_qty):
    if status == "Krytyczne":
        return f"Uruchom produkcję natychmiast, minimum {shortage_qty:,.0f} szt."
    if status == "Wysokie ryzyko":
        return f"Zaplanuj produkcję w najbliższym oknie, minimum {shortage_qty:,.0f} szt."
    if status == "Ryzyko":
        return "Zabezpiecz slot produkcyjny i monitoruj najbliższe wysyłki."
    if status == "Monitoruj":
        return "Monitoruj zapas i przygotuj plan uzupełnienia."
    if status == "Pokryte":
        return "Brak pilnej akcji. Zapotrzebowanie jest pokryte."
    return "Brak zapotrzebowania w wybranym zakresie."


def calculate_planner_outputs(planner_source, planner_input_df, today=None):
    today = pd.Timestamp(today).normalize() if today is not None else pd.Timestamp.now().normalize()
    if planner_source is None or planner_source.empty:
        empty_summary = pd.DataFrame(
            columns=[
                "Part Number",
                "Part Description",
                "Stock",
                "Safety Stock",
                "Total Demand",
                "Coverage Until",
                "First Shortage Date",
                "Days Covered",
                "Shortage Qty",
                "Qty To Produce Now",
                "Coverage %",
                "Status",
                "Recommendation",
                "Production Priority",
            ]
        )
        empty_daily = pd.DataFrame(
            columns=[
                "Part Number",
                "Part Description",
                "Ship Date",
                "Demand Qty",
                "Stock",
                "Safety Stock",
                "Available Stock",
                "Cumulative Demand",
                "Remaining Stock",
                "Remaining Above Safety",
                "Shortage On Day",
                "Shortage Flag",
            ]
        )
        return empty_summary, empty_daily

    inputs = planner_input_df.copy()
    inputs["Stock"] = pd.to_numeric(inputs["Stock"], errors="coerce").fillna(0.0)
    inputs["Safety Stock"] = pd.to_numeric(inputs["Safety Stock"], errors="coerce").fillna(0.0)

    daily_frames = []
    summary_rows = []

    for input_row in inputs.to_dict("records"):
        part_number = input_row["Part Number"]
        product_source = planner_source[planner_source["Part Number"] == part_number].copy()
        if product_source.empty:
            continue

        product_source = product_source.sort_values("Ship Date").reset_index(drop=True)
        stock = float(input_row.get("Stock", 0.0))
        safety_stock = float(input_row.get("Safety Stock", 0.0))
        available_stock = stock - safety_stock

        product_source["Stock"] = stock
        product_source["Safety Stock"] = safety_stock
        product_source["Available Stock"] = available_stock
        product_source["Cumulative Demand"] = product_source["Demand Qty"].cumsum()
        product_source["Remaining Stock"] = stock - product_source["Cumulative Demand"]
        product_source["Remaining Above Safety"] = available_stock - product_source["Cumulative Demand"]
        product_source["Shortage On Day"] = product_source["Remaining Above Safety"].apply(
            lambda value: max(abs(value), 0.0) if value < 0 else 0.0
        )
        product_source["Shortage Flag"] = product_source["Remaining Above Safety"] < 0
        daily_frames.append(product_source)

        total_demand = float(product_source["Demand Qty"].sum())
        shortage_rows = product_source[product_source["Shortage Flag"]]
        first_shortage_date = (
            shortage_rows["Ship Date"].iloc[0] if not shortage_rows.empty else pd.NaT
        )
        covered_rows = product_source[~product_source["Shortage Flag"]]
        coverage_until = (
            covered_rows["Ship Date"].iloc[-1]
            if not covered_rows.empty
            else pd.NaT
        )
        if shortage_rows.empty and not product_source.empty:
            coverage_until = product_source["Ship Date"].iloc[-1]

        days_covered = 0
        coverage_anchor = coverage_until if not pd.isna(coverage_until) else first_shortage_date
        if not pd.isna(coverage_anchor):
            days_covered = max(int((pd.Timestamp(coverage_anchor).normalize() - today).days), 0)

        shortage_qty = max(total_demand + safety_stock - stock, 0.0)
        qty_to_produce_now = shortage_qty
        denominator = total_demand + safety_stock
        coverage_pct = 100.0 if denominator <= 0 else (stock / denominator) * 100.0
        status = _classify_status(first_shortage_date, shortage_qty, coverage_pct, total_demand, today)

        summary_rows.append(
            {
                "Part Number": part_number,
                "Part Description": input_row.get("Part Description", ""),
                "Stock": stock,
                "Safety Stock": safety_stock,
                "Total Demand": total_demand,
                "Coverage Until": coverage_until,
                "First Shortage Date": first_shortage_date,
                "Days Covered": days_covered,
                "Shortage Qty": shortage_qty,
                "Qty To Produce Now": qty_to_produce_now,
                "Coverage %": coverage_pct,
                "Status": status,
                "Recommendation": _build_recommendation(status, qty_to_produce_now),
                "_status_rank": _status_rank(status),
            }
        )

    planner_daily = pd.concat(daily_frames, ignore_index=True) if daily_frames else pd.DataFrame()
    planner_results = pd.DataFrame(summary_rows)
    if planner_results.empty:
        return calculate_planner_outputs(_empty_planner_source(), pd.DataFrame(), today=today)

    planner_results["_shortage_sort"] = planner_results["First Shortage Date"].fillna(pd.Timestamp.max)
    planner_results = planner_results.sort_values(
        ["_status_rank", "_shortage_sort", "Shortage Qty", "Part Number"],
        ascending=[True, True, False, True],
    ).reset_index(drop=True)
    planner_results["Production Priority"] = range(1, len(planner_results) + 1)
    planner_results = planner_results.drop(columns=["_status_rank", "_shortage_sort"])
    return planner_results, planner_daily


def build_planner_kpis(planner_results):
    if planner_results is None or planner_results.empty:
        return {
            "products": 0,
            "critical": 0,
            "to_produce": 0.0,
            "covered_share": 0.0,
            "avg_coverage": 0.0,
        }

    critical = planner_results["Status"].isin(["Krytyczne", "Wysokie ryzyko"]).sum()
    covered = planner_results["Status"].eq("Pokryte").sum()
    return {
        "products": int(len(planner_results)),
        "critical": int(critical),
        "to_produce": float(planner_results["Qty To Produce Now"].sum()),
        "covered_share": float((covered / len(planner_results)) * 100.0),
        "avg_coverage": float(planner_results["Coverage %"].mean()),
    }


def build_planner_priority_chart(planner_results):
    if planner_results is None or planner_results.empty:
        return None

    chart_source = planner_results.head(12).copy()
    chart_source["Display Label"] = chart_source["Part Number"] + " | " + chart_source["Part Description"].str.slice(0, 28)
    chart_source["Status Color"] = chart_source["Status"].map(
        {
            "Krytyczne": "#ff6b6b",
            "Wysokie ryzyko": "#ff9f43",
            "Ryzyko": "#f6c453",
            "Monitoruj": "#60a5fa",
            "Pokryte": "#34d399",
            "Brak popytu": "#94a3b8",
        }
    ).fillna("#94a3b8")

    bars = (
        alt.Chart(chart_source)
        .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6, opacity=0.92)
        .encode(
            x=alt.X("Qty To Produce Now:Q", title="Qty To Produce Now"),
            y=alt.Y("Display Label:N", sort="-x", title=None),
            color=alt.Color("Status Color:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("Production Priority:Q", title="Priority"),
                alt.Tooltip("Part Number:N"),
                alt.Tooltip("Part Description:N"),
                alt.Tooltip("Qty To Produce Now:Q", format=",.0f"),
                alt.Tooltip("First Shortage Date:T", title="First Shortage"),
                alt.Tooltip("Status:N"),
            ],
        )
    )
    labels = (
        alt.Chart(chart_source)
        .mark_text(align="left", dx=8, color="#f8fafc", fontWeight="bold")
        .encode(
            x=alt.X("Qty To Produce Now:Q"),
            y=alt.Y("Display Label:N", sort="-x", title=None),
            text=alt.Text("Qty To Produce Now:Q", format=",.0f"),
        )
    )
    return bars + labels


def build_planner_coverage_chart(planner_results):
    if planner_results is None or planner_results.empty:
        return None

    chart_source = planner_results.copy()
    chart_source["Display Label"] = chart_source["Part Number"] + " | " + chart_source["Part Description"].str.slice(0, 24)
    chart_source["Status Color"] = chart_source["Status"].map(
        {
            "Krytyczne": "#ff6b6b",
            "Wysokie ryzyko": "#ff9f43",
            "Ryzyko": "#f6c453",
            "Monitoruj": "#60a5fa",
            "Pokryte": "#34d399",
            "Brak popytu": "#94a3b8",
        }
    ).fillna("#94a3b8")
    threshold_source = pd.DataFrame({"threshold": [100.0]})

    bars = (
        alt.Chart(chart_source)
        .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, opacity=0.88)
        .encode(
            x=alt.X("Display Label:N", sort=alt.EncodingSortField("Production Priority", order="ascending"), title=None, axis=alt.Axis(labelAngle=-28)),
            y=alt.Y("Coverage %:Q", title="Coverage %"),
            color=alt.Color("Status Color:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("Part Number:N"),
                alt.Tooltip("Coverage %:Q", format=".1f"),
                alt.Tooltip("Stock:Q", format=",.0f"),
                alt.Tooltip("Safety Stock:Q", format=",.0f"),
                alt.Tooltip("Total Demand:Q", format=",.0f"),
                alt.Tooltip("Status:N"),
            ],
        )
    )
    rule = alt.Chart(threshold_source).mark_rule(color="#e2e8f0", strokeDash=[4, 4]).encode(y="threshold:Q")
    return bars + rule


def build_planner_display_table(planner_results):
    if planner_results is None or planner_results.empty:
        return pd.DataFrame()

    table = planner_results.copy()
    for column in ["Coverage Until", "First Shortage Date"]:
        table[column] = pd.to_datetime(table[column], errors="coerce").dt.strftime("%Y-%m-%d")
        table[column] = table[column].fillna("n/a")

    for column in ["Stock", "Safety Stock", "Total Demand", "Shortage Qty", "Qty To Produce Now"]:
        table[column] = table[column].map(lambda value: f"{float(value):,.0f}")
    table["Coverage %"] = table["Coverage %"].map(lambda value: f"{float(value):.1f}%")
    return table


def build_planner_daily_display(planner_daily, part_number):
    if planner_daily is None or planner_daily.empty:
        return pd.DataFrame()

    detail = planner_daily[planner_daily["Part Number"] == part_number].copy()
    if detail.empty:
        return pd.DataFrame()

    detail["Ship Date"] = pd.to_datetime(detail["Ship Date"]).dt.strftime("%Y-%m-%d")
    for column in [
        "Demand Qty",
        "Stock",
        "Safety Stock",
        "Available Stock",
        "Cumulative Demand",
        "Remaining Stock",
        "Remaining Above Safety",
        "Shortage On Day",
    ]:
        detail[column] = detail[column].map(lambda value: f"{float(value):,.0f}")
    detail["Shortage Flag"] = detail["Shortage Flag"].map(lambda value: "Tak" if value else "Nie")
    return detail


def build_planner_excel_bytes(planner_inputs, planner_results, planner_daily):
    output = io.BytesIO()
    planner_inputs_export = planner_inputs.copy()
    planner_results_export = build_planner_display_table(planner_results)
    planner_daily_export = planner_daily.copy()
    if not planner_daily_export.empty:
        planner_daily_export["Ship Date"] = pd.to_datetime(planner_daily_export["Ship Date"]).dt.strftime("%Y-%m-%d")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        planner_inputs_export.to_excel(writer, sheet_name="Planner Inputs", index=False)
        planner_results_export.to_excel(writer, sheet_name="Planner Summary", index=False)
        planner_daily_export.to_excel(writer, sheet_name="Planner Daily", index=False)

    return output.getvalue()
