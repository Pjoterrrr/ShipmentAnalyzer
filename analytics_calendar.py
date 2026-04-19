from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta

import pandas as pd


DAY_NAME_PL = {
    0: "Poniedzialek",
    1: "Wtorek",
    2: "Sroda",
    3: "Czwartek",
    4: "Piatek",
    5: "Sobota",
    6: "Niedziela",
}


@dataclass(frozen=True)
class ReferenceWeek:
    reference_date: date
    week_start: date
    week_end: date
    iso_year: int
    iso_week: int
    week_label: str


def normalize_to_date(value: date | pd.Timestamp | str) -> date:
    if isinstance(value, date) and not isinstance(value, pd.Timestamp):
        return value
    return pd.Timestamp(value).date()


def week_label_for_date(value: date | pd.Timestamp | str) -> str:
    normalized = normalize_to_date(value)
    iso_year, iso_week, _ = normalized.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


def week_bounds_for_date(value: date | pd.Timestamp | str) -> tuple[date, date]:
    normalized = normalize_to_date(value)
    week_start = normalized - timedelta(days=normalized.weekday())
    return week_start, week_start + timedelta(days=6)


def get_last_completed_reference_week(reference_date: date | pd.Timestamp | str) -> ReferenceWeek:
    normalized = normalize_to_date(reference_date)
    week_offset = 0 if normalized.weekday() == 6 else normalized.weekday() + 1
    week_end = normalized - timedelta(days=week_offset)
    week_start = week_end - timedelta(days=6)
    iso_year, iso_week, _ = week_end.isocalendar()
    return ReferenceWeek(
        reference_date=normalized,
        week_start=week_start,
        week_end=week_end,
        iso_year=iso_year,
        iso_week=iso_week,
        week_label=f"{iso_year}-W{iso_week:02d}",
    )


def add_iso_week_columns(frame: pd.DataFrame, date_column: str) -> pd.DataFrame:
    enriched = frame.copy()
    enriched[date_column] = pd.to_datetime(enriched[date_column], errors="coerce")
    iso_calendar = enriched[date_column].dt.isocalendar()
    enriched["ISO Year"] = iso_calendar["year"].astype("Int64")
    enriched["ISO Week"] = iso_calendar["week"].astype("Int64")
    enriched["Week Start"] = enriched[date_column] - pd.to_timedelta(
        enriched[date_column].dt.weekday, unit="D"
    )
    enriched["Week End"] = enriched["Week Start"] + pd.Timedelta(days=6)
    enriched["Week Label"] = enriched.apply(
        lambda row: (
            f"{int(row['ISO Year'])}-W{int(row['ISO Week']):02d}"
            if pd.notna(row["ISO Year"]) and pd.notna(row["ISO Week"])
            else pd.NA
        ),
        axis=1,
    )
    return enriched


def easter_sunday(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def polish_public_holidays(year: int) -> dict[date, str]:
    easter = easter_sunday(year)
    return {
        date(year, 1, 1): "Nowy Rok",
        date(year, 1, 6): "Trzech Kroli",
        easter: "Wielkanoc",
        easter + timedelta(days=1): "Poniedzialek Wielkanocny",
        date(year, 5, 1): "Swieto Pracy",
        date(year, 5, 3): "Swieto Konstytucji 3 Maja",
        easter + timedelta(days=49): "Zeslanie Ducha Swietego",
        easter + timedelta(days=60): "Boze Cialo",
        date(year, 8, 15): "Wniebowziecie NMP",
        date(year, 11, 1): "Wszystkich Swietych",
        date(year, 11, 11): "Narodowe Swieto Niepodleglosci",
        date(year, 12, 25): "Boze Narodzenie",
        date(year, 12, 26): "Drugi Dzien Bozego Narodzenia",
    }


def classify_polish_day(value: date | pd.Timestamp | str) -> dict[str, object]:
    normalized = normalize_to_date(value)
    holiday_name = polish_public_holidays(normalized.year).get(normalized)
    is_weekend = normalized.weekday() >= 5
    is_working_day = normalized.weekday() < 5 and holiday_name is None
    if holiday_name:
        day_type = "Holiday"
    elif normalized.weekday() == 5:
        day_type = "Saturday"
    elif normalized.weekday() == 6:
        day_type = "Sunday"
    else:
        day_type = "Working Day"
    return {
        "Date": normalized,
        "Day Name": DAY_NAME_PL[normalized.weekday()],
        "Holiday Name": holiday_name or "",
        "Day Type": day_type,
        "Is Weekend": is_weekend,
        "Is Holiday": holiday_name is not None,
        "Is Working Day": is_working_day,
    }


def build_calendar_frame(start_date: date | pd.Timestamp | str, end_date: date | pd.Timestamp | str) -> pd.DataFrame:
    start = normalize_to_date(start_date)
    end = normalize_to_date(end_date)
    if end < start:
        start, end = end, start
    calendar = pd.DataFrame({"Date": pd.date_range(start, end, freq="D")})
    calendar["Date"] = calendar["Date"].dt.date
    calendar_details = calendar["Date"].map(classify_polish_day).apply(pd.Series)
    calendar = pd.concat([calendar.drop(columns=["Date"]), calendar_details], axis=1)
    calendar = add_iso_week_columns(calendar, "Date")
    calendar["Week Start"] = pd.to_datetime(calendar["Week Start"]).dt.date
    calendar["Week End"] = pd.to_datetime(calendar["Week End"]).dt.date
    calendar["ISO Week Label"] = calendar["Week Label"]
    return calendar[
        [
            "Date",
            "Day Name",
            "Day Type",
            "Holiday Name",
            "Is Weekend",
            "Is Holiday",
            "Is Working Day",
            "ISO Year",
            "ISO Week",
            "ISO Week Label",
            "Week Start",
            "Week End",
        ]
    ]


def safe_percent_change(current_value: float, previous_value: float) -> float | None:
    current = float(current_value)
    previous = float(previous_value)
    if previous == 0:
        return 0.0 if current == 0 else None
    return round(((current - previous) / previous) * 100, 2)


def format_percent_change(current_value: float, previous_value: float) -> str:
    percent_value = safe_percent_change(current_value, previous_value)
    if percent_value is None:
        return "new"
    return f"{percent_value:+.2f}%"


def is_change_alert(current_value: float, previous_value: float, threshold: float) -> bool:
    percent_value = safe_percent_change(current_value, previous_value)
    if percent_value is None:
        return float(current_value) != 0
    return abs(percent_value) >= float(threshold)


def build_weekly_summary(
    frame: pd.DataFrame,
    date_column: str,
    range_start: date | pd.Timestamp | str,
    range_end: date | pd.Timestamp | str,
    reference_date: date | pd.Timestamp | str,
    threshold: float,
) -> pd.DataFrame:
    range_start_value = normalize_to_date(range_start)
    range_end_value = normalize_to_date(range_end)
    if range_end_value < range_start_value:
        range_start_value, range_end_value = range_end_value, range_start_value

    reference_week = get_last_completed_reference_week(reference_date)
    calendar = build_calendar_frame(range_start_value, range_end_value)
    weeks_in_scope = (
        calendar.groupby(
            ["ISO Year", "ISO Week", "ISO Week Label", "Week Start", "Week End"],
            as_index=False,
        )
        .agg(
            Working_Days_PL=("Is Working Day", "sum"),
            Range_Days=("Date", "count"),
            Holidays_PL=("Is Holiday", "sum"),
            Weekend_Days=("Is Weekend", "sum"),
        )
        .rename(columns={"ISO Week Label": "Week Label"})
        .sort_values(["Week Start", "Week End"])
    )

    if frame.empty:
        weekly = weeks_in_scope.copy()
        weekly["Quantity_Prev"] = 0.0
        weekly["Quantity_Curr"] = 0.0
        weekly["Delta"] = 0.0
        weekly["Alert Rows"] = 0
        weekly["Products"] = 0
    else:
        weekly_source = add_iso_week_columns(frame, date_column)
        weekly_source = weekly_source.dropna(subset=["Week Label"]).copy()
        weekly = (
            weekly_source.groupby(
                ["ISO Year", "ISO Week", "Week Label", "Week Start", "Week End"],
                as_index=False,
            )
            .agg(
                Quantity_Prev=("Quantity_Prev", "sum"),
                Quantity_Curr=("Quantity_Curr", "sum"),
                Delta=("Delta", "sum"),
                **{
                    "Alert Rows": ("Alert", "sum"),
                    "Products": ("Product Label", "nunique"),
                },
            )
            .sort_values(["Week Start", "Week End"])
        )
        weekly["Week Start"] = pd.to_datetime(weekly["Week Start"]).dt.date
        weekly["Week End"] = pd.to_datetime(weekly["Week End"]).dt.date
        weekly = weeks_in_scope.merge(
            weekly,
            on=["ISO Year", "ISO Week", "Week Label", "Week Start", "Week End"],
            how="left",
        )
        weekly[["Quantity_Prev", "Quantity_Curr", "Delta"]] = weekly[
            ["Quantity_Prev", "Quantity_Curr", "Delta"]
        ].fillna(0.0)
        weekly["Alert Rows"] = weekly["Alert Rows"].fillna(0).astype(int)
        weekly["Products"] = weekly["Products"].fillna(0).astype(int)

    weekly = weekly.sort_values(["Week Start", "Week End"]).reset_index(drop=True)
    weekly["Week Label Short"] = weekly["ISO Week"].map(lambda value: f"Week {int(value)}")
    weekly["Is Partial Range"] = weekly["Range_Days"] < 7
    weekly["Is Closed Week"] = weekly["Week End"].map(lambda value: value <= reference_week.week_end)
    weekly["Is Reference Week"] = weekly["Week Label"].eq(reference_week.week_label)
    weekly["Week Status"] = weekly.apply(
        lambda row: (
            "Open week"
            if not row["Is Closed Week"]
            else "Partial range"
            if row["Is Partial Range"]
            else "Closed full week"
        ),
        axis=1,
    )
    weekly["Avg Current / Working Day"] = weekly.apply(
        lambda row: round(row["Quantity_Curr"] / row["Working_Days_PL"], 2)
        if row["Working_Days_PL"] > 0
        else pd.NA,
        axis=1,
    )
    weekly["Avg Previous / Working Day"] = weekly.apply(
        lambda row: round(row["Quantity_Prev"] / row["Working_Days_PL"], 2)
        if row["Working_Days_PL"] > 0
        else pd.NA,
        axis=1,
    )
    weekly["Release Percent Change"] = weekly.apply(
        lambda row: safe_percent_change(row["Quantity_Curr"], row["Quantity_Prev"]),
        axis=1,
    )
    weekly["Release Percent Label"] = weekly.apply(
        lambda row: format_percent_change(row["Quantity_Curr"], row["Quantity_Prev"]),
        axis=1,
    )
    weekly["Release Alert"] = weekly.apply(
        lambda row: is_change_alert(row["Quantity_Curr"], row["Quantity_Prev"], threshold),
        axis=1,
    )
    weekly["Previous Week Current Qty"] = weekly["Quantity_Curr"].shift(1).fillna(0.0)
    weekly["WoW Delta"] = weekly["Quantity_Curr"] - weekly["Previous Week Current Qty"]
    weekly["WoW Percent Change"] = weekly.apply(
        lambda row: safe_percent_change(row["Quantity_Curr"], row["Previous Week Current Qty"]),
        axis=1,
    )
    weekly["WoW Percent Label"] = weekly.apply(
        lambda row: format_percent_change(row["Quantity_Curr"], row["Previous Week Current Qty"]),
        axis=1,
    )
    weekly["WoW Alert"] = weekly.apply(
        lambda row: is_change_alert(row["Quantity_Curr"], row["Previous Week Current Qty"], threshold),
        axis=1,
    )
    weekly["Any Weekly Alert"] = weekly["Release Alert"] | weekly["WoW Alert"]
    return weekly

