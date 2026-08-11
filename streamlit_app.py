import re
from datetime import date, datetime
from html import escape
from pathlib import Path
from typing import List, Tuple
import subprocess
import zoneinfo

import streamlit as st
import pandas as pd
import plotly.express as px
from openpyxl import load_workbook
from google import genai

WEEK_SHEET_RE = re.compile(r"^Week\s+(\d{1,2})-(\d{4})$", re.IGNORECASE)
MASTER_WORKBOOK = Path(__file__).with_name("master.xlsx")
DEFAULT_PLAYER = "Ela Velic"
LOWER_IS_BETTER = {"Rank", "Singles WTN"}
PERCENT_METRICS = {"Singles W/L Career %", "Singles W/L-YTD %"}
DISPLAY_COLUMNS = [
    "Week Label", "Rank", "Points", "Tournaments",
    "Singles WTN", "Career Matches", "Singles W/L-Career",
    "Singles W/L-YTD", "Singles W/L Career %", "Singles W/L-YTD %",
]
LATEST_TRENDS_COLUMNS = ["Rank", "Points", "Tournaments", "Singles WTN", "Career Matches", "Singles W/L Career %", "Singles W/L-YTD %"]
NUMERIC_TREND_COLUMNS = LATEST_TRENDS_COLUMNS.copy()
DEFAULT_CHART_METRICS = ["Rank", "Points", "Singles WTN"]
OPTIONAL_CHART_METRICS = ["Rank", "Points", "Singles WTN", "Tournaments", "Career Matches", "Singles W/L Career %", "Singles W/L-YTD %"]
YTD_TREND_METRICS = LATEST_TRENDS_COLUMNS.copy()
TOP_MOVER_METRIC_COLUMNS = ["Points", "Singles WTN", "Tournaments", "Career Matches", "Singles W/L-YTD %"]


def strip_leading_zero(value: str) -> str:
    return value.replace(" 0", " ")


def format_custom_datetime(dt: datetime) -> str:
    month_day_year = dt.strftime("%B %d, %Y")
    time_str = dt.strftime("%I:%M%p").lower().lstrip("0")
    return f"{month_day_year} - {time_str}"


def get_data_updated_text(workbook_path: Path) -> str:
    try:
        git_epoch_str = subprocess.check_output(
            ["git", "log", "-1", "--format=%at", "master.xlsx"]
        ).decode("utf-8").strip()
        if git_epoch_str:
            epoch_time = float(git_epoch_str)
            try:
                est_tz = zoneinfo.ZoneInfo("America/New_York")
                dt = datetime.fromtimestamp(epoch_time, est_tz)
            except Exception:
                dt = datetime.fromtimestamp(epoch_time)
            return format_custom_datetime(dt)
    except Exception:
        pass
    
    if workbook_path.exists():
        try:
            est_tz = zoneinfo.ZoneInfo("America/New_York")
            dt = datetime.fromtimestamp(workbook_path.stat().st_mtime, est_tz)
        except Exception:
            dt = datetime.fromtimestamp(workbook_path.stat().st_mtime)
        return format_custom_datetime(dt)
    return "Unavailable"


def format_week_date_range(year: int, week_number: int) -> str:
    start = date.fromisocalendar(year, week_number, 1)
    end = date.fromisocalendar(year, week_number, 7)
    if start.year == end.year and start.month == end.month:
        return f"{start.strftime('%B')} {start.day} to {end.strftime('%B')} {end.day}, {end.year}"
    return f"{start.strftime('%B')} {start.day}, {start.year} to {end.strftime('%B')} {end.day}, {end.year}"


def latest_dataset_text(rankings: pd.DataFrame, workbook_path: Path) -> Tuple[str, str]:
    latest_sort = rankings["Week Sort"].max()
    latest_row = rankings[rankings["Week Sort"] == latest_sort].iloc[0]
    week_number = int(latest_row["Week Number"])
    ranking_year = int(latest_row["Ranking Year"])
    range_text = format_week_date_range(ranking_year, week_number)
    return (
        f"Latest dataset: Week {week_number}, {ranking_year} ({range_text})",
        f"Data updated: {get_data_updated_text(workbook_path)}",
    )


def parse_week_sheet_name(sheet_name: str) -> Tuple[int, int, str]:
    match = WEEK_SHEET_RE.match(str(sheet_name).strip())
    if not match:
        raise ValueError(f"Not a weekly ranking sheet: {sheet_name}")
    week_number = int(match.group(1))
    ranking_year = int(match.group(2))
    return week_number, ranking_year, f"Week {week_number:02d}-{ranking_year}"


def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def extract_hyperlink_formula_url(value):
    if not isinstance(value, str):
        return None
    formula = value.strip()
    if not formula.upper().startswith("=HYPERLINK("):
        return None
    match = re.search(r'=HYPERLINK\(\s*["\']([^"\']+)["\']', formula, flags=re.IGNORECASE)
    return match.group(1).strip() if match else None


def normalize_external_url(value):
    if not isinstance(value, str):
        return None
    url = value.strip()
    return url if url.lower().startswith(("http://", "https://")) else None


def coerce_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    numeric_cols = ["Rank", "Year of Birth", "Points", "Tournaments", "Singles WTN", "Career Matches", "Singles W/L Career %", "Singles W/L-YTD %"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].replace({"unavailable": pd.NA, "Unavailable": pd.NA, "UNAVAILABLE": pd.NA}).pipe(pd.to_numeric, errors="coerce")
    return df


def extract_profile_urls(ws, df: pd.DataFrame) -> List[object]:
    header_lookup = {str(cell.value).strip(): cell.column for cell in ws[1] if cell.value is not None}
    profile_col = header_lookup.get("Profile")
    if profile_col is None:
        return [pd.NA] * len(df)
    urls: List[object] = []
    for row_index in df.index:
        cell = ws.cell(row=int(row_index) + 2, column=profile_col)
        candidates = []
        if cell.hyperlink is not None:
            candidates.extend([cell.hyperlink.target, cell.hyperlink.location])
        candidates.extend([extract_hyperlink_formula_url(cell.value), cell.value])
        external_url = next((normalize_external_url(candidate) for candidate in candidates if normalize_external_url(candidate)), None)
        urls.append(external_url if external_url else pd.NA)
    return urls


def get_latest_profile_url(player_df: pd.DataFrame):
    if "Profile URL" not in player_df.columns:
        return None
    for url in player_df.sort_values("Week Sort", ascending=False)["Profile URL"].tolist():
        external_url = normalize_external_url(url)
        if external_url:
            return external_url
    return None


@st.cache_data(show_spinner=False)
def load_rankings(excel_file: Path) -> pd.DataFrame:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, engine="openpyxl")
    workbook = load_workbook(excel_file, data_only=False, read_only=False)
    frames: List[pd.DataFrame] = []
    for sheet_name, sheet_df in all_sheets.items():
        sheet_key = str(sheet_name).strip().lower()
        if sheet_key == "trends" or not WEEK_SHEET_RE.match(str(sheet_name).strip()):
            continue
        week_number, ranking_year, week_label = parse_week_sheet_name(sheet_name)
        df = clean_headers(sheet_df)
        if "Player" not in df.columns:
            continue
        df["Profile URL"] = extract_profile_urls(workbook[sheet_name], df) if sheet_name in workbook.sheetnames else pd.NA
        df = df.dropna(how="all")
        df = df[df["Player"].notna()]
        df["Player"] = df["Player"].astype(str).str.strip()
        df = df[df["Player"] != ""]
        df["Week Number"] = week_number
        df["Ranking Year"] = ranking_year
        df["Week Label"] = week_label
        df["Week Sort"] = ranking_year * 100 + week_number
        frames.append(df)
    workbook.close()
    if not frames:
        raise ValueError("No weekly sheets found. Expected sheet names like 'Week 27-2026'.")
    return coerce_numeric_columns(pd.concat(frames, ignore_index=True, sort=False)).sort_values(["Week Sort", "Rank", "Player"], na_position="last")


def format_value(metric: str, value) -> str:
    if pd.isna(value):
        return "unavailable"
    if metric in PERCENT_METRICS:
        return f"{value:.0%}"
    if metric == "Year of Birth":
        return f"{int(round(value))}"
    if metric in {"Rank", "Ontario Rank", "Tournaments", "Career Matches"}:
        return f"{int(round(value)):,}"
    if metric == "Singles WTN":
        return f"{value:.1f}"
    if metric == "Points":
        return f"{value:,.3f}"
    return f"{value:,.2f}"


def performance_change(metric: str, previous_value, current_value):
    if pd.isna(previous_value) or pd.isna(current_value):
        return pd.NA
    return previous_value - current_value if metric in LOWER_IS_BETTER else current_value - previous_value


def format_trend(metric: str, change) -> str:
    if pd.isna(change):
        return "—"
    if abs(float(change)) < 1e-12:
        return "→ 0"
    arrow = "↑" if change > 0 else "↓"
    sign = "+" if change > 0 else ""
    if metric in PERCENT_METRICS:
        return f"{arrow} {sign}{change:.0%} pp"
    if metric in {"Rank", "Tournaments", "Career Matches"}:
        return f"{arrow} {sign}{change:.0f}"
    if metric == "Singles WTN":
        return f"{arrow} {sign}{change:.1f}"
    if metric == "Points":
        return f"{arrow} {sign}{change:,.3f}"
    return f"{arrow} {sign}{change:,.2f}"


def format_delta_for_metric_card(metric: str, change):
    if pd.isna(change):
        return None
    if abs(float(change)) < 1e-12:
        return "0"
    sign = "+" if change > 0 else ""
    if metric in PERCENT_METRICS:
        return f"{sign}{change:.0%} pp"
    if metric in {"Rank", "Tournaments", "Career Matches"}:
        return f"{sign}{change:.0f}"
    if metric == "Singles WTN":
        return f"{sign}{change:.1f}"
    if metric == "Points":
        return f"{sign}{change:,.3f}"
    return f"{sign}{change:,.2f}"


def value_with_week_trend(metric: str, previous_value, current_value) -> str:
    return f"{format_value(metric, current_value)} ({format_trend(metric, performance_change(metric, previous_value, current_value))})"


def color_trend_cell(value: object) -> str:
    text = str(value)
    if "↑" in text:
        return "color: #0f7b3f; font-weight: 600;"
    if "↓" in text:
        return "color: #b42318; font-weight: 600;"
    if "→" in text:
        return "color: #6b7280;"
    return ""


def style_trend_table(display_df: pd.DataFrame, trend_cols: List[str]):
    cols = [c for c in trend_cols if c in display_df.columns]
    styler = display_df.style
    if not cols:
        return styler
    if hasattr(styler, "map"):
        return styler.map(color_trend_cell, subset=cols)
    return styler.applymap(color_trend_cell, subset=cols)


def table_auto_height(row_count: int, row_height: int = 35, header_height: int = 42, padding: int = 8) -> int:
    return int(header_height + padding + max(row_count, 1) * row_height)


def get_latest_provincial_rank(rankings: pd.DataFrame, selected_player: str):
    latest_sort = rankings["Week Sort"].max()
    latest_df = rankings[rankings["Week Sort"] == latest_sort].copy()
    player_rows = latest_df[latest_df["Player"] == selected_player]
    if player_rows.empty or "Province" not in latest_df.columns:
        return pd.NA
    player_province = player_rows.iloc[0].get("Province")
    if pd.isna(player_province):
        return pd.NA
    province_df = latest_df[latest_df["Province"].astype(str).str.strip() == str(player_province).strip()]
    province_df = province_df.sort_values(["Rank", "Player"], na_position="last").reset_index(drop=True)
    matches = province_df[province_df["Player"] == selected_player]
    if matches.empty:
        return pd.NA
    return int(matches.index[0] + 1)


def build_player_history_display(player_df: pd.DataFrame) -> pd.DataFrame:
    chronological_df = player_df.sort_values("Week Sort").copy()
    excluded_cols = {"Province", "Club"}
    visible_cols = [c for c in DISPLAY_COLUMNS if c in chronological_df.columns and c not in excluded_cols]
    display_df = chronological_df[visible_cols + ["Week Sort"]].copy()
    for metric in NUMERIC_TREND_COLUMNS:
        if metric not in chronological_df.columns or metric not in display_df.columns:
            continue
        previous_values = chronological_df[metric].shift(1)
        display_df[metric] = [format_value(metric, v) if i == 0 else value_with_week_trend(metric, previous_values.iloc[i], v) for i, v in enumerate(chronological_df[metric].tolist())]
    display_df = display_df.sort_values("Week Sort", ascending=False).drop(columns=["Week Sort"])
    return display_df.rename(columns={"Week Label": "Week"}) if "Week Label" in display_df.columns else display_df


def build_latest_week_trends(rankings: pd.DataFrame) -> Tuple[pd.DataFrame, str, str]:
    week_sorts = sorted(rankings["Week Sort"].dropna().unique())
    if len(week_sorts) < 2:
        raise ValueError("At least two weekly sheets are needed to calculate latest week trends.")
    previous_sort, latest_sort = week_sorts[-2], week_sorts[-1]
    previous_label = rankings.loc[rankings["Week Sort"] == previous_sort, "Week Label"].iloc[0]
    latest_label = rankings.loc[rankings["Week Sort"] == latest_sort, "Week Label"].iloc[0]
    latest_df = rankings[rankings["Week Sort"] == latest_sort].copy().sort_values("Rank", na_position="last")
    previous_by_player = rankings[rankings["Week Sort"] == previous_sort].drop_duplicates("Player").set_index("Player")
    base_cols = [c for c in ["Rank", "Player", "Year of Birth", "Province", "Club"] if c in latest_df.columns]
    display_df = latest_df[base_cols].copy()
    if "Year of Birth" in display_df.columns:
        display_df["Year of Birth"] = display_df["Year of Birth"].apply(lambda v: format_value("Year of Birth", v))
    for metric in LATEST_TRENDS_COLUMNS:
        if metric not in latest_df.columns:
            continue
        cells = []
        for _, latest_row in latest_df.iterrows():
            player = latest_row.get("Player")
            previous_value = previous_by_player.loc[player, metric] if player in previous_by_player.index and metric in previous_by_player.columns else pd.NA
            cells.append(value_with_week_trend(metric, previous_value, latest_row.get(metric)))
        display_df[metric] = cells
    return display_df, previous_label, latest_label


def build_top_movers(rankings: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, str, str]:
    week_sorts = sorted(rankings["Week Sort"].dropna().unique())
    if len(week_sorts) < 2:
        raise ValueError("At least two weekly sheets are needed to calculate top movers.")
    previous_sort, latest_sort = week_sorts[-2], week_sorts[-1]
    previous_label = rankings.loc[rankings["Week Sort"] == previous_sort, "Week Label"].iloc[0]
    latest_label = rankings.loc[rankings["Week Sort"] == latest_sort, "Week Label"].iloc[0]
    latest_df = rankings[rankings["Week Sort"] == latest_sort].drop_duplicates("Player").copy()
    previous_df = rankings[rankings["Week Sort"] == previous_sort].drop_duplicates("Player").copy()
    keep_cols = ["Player", "Rank"] + [c for c in TOP_MOVER_METRIC_COLUMNS if c in rankings.columns]
    merged = latest_df[keep_cols].merge(previous_df[keep_cols], on="Player", suffixes=(" Latest", " Previous"), how="inner")
    merged["Rank Move"] = merged.apply(lambda row: performance_change("Rank", row["Rank Previous"], row["Rank Latest"]), axis=1)
    merged = merged.dropna(subset=["Rank Move"])
    merged = merged[merged["Rank Move"] != 0]
    def make_table(source: pd.DataFrame, ascending: bool) -> pd.DataFrame:
        if source.empty:
            return pd.DataFrame(columns=["Player", f"Previous Rank ({previous_label})", f"Latest Rank ({latest_label})", "Rank Trend"])
        source = source.sort_values("Rank Move", ascending=ascending).head(10)
        rows = []
        for _, row in source.iterrows():
            out = {"Player": row["Player"], f"Previous Rank ({previous_label})": format_value("Rank", row["Rank Previous"]), f"Latest Rank ({latest_label})": format_value("Rank", row["Rank Latest"]), "Rank Trend": format_trend("Rank", row["Rank Move"])}
            for metric in TOP_MOVER_METRIC_COLUMNS:
                latest_col, previous_col = f"{metric} Latest", f"{metric} Previous"
                if latest_col in row.index and previous_col in row.index:
                    out[metric] = value_with_week_trend(metric, row[previous_col], row[latest_col])
            rows.append(out)
        return pd.DataFrame(rows)
    return make_table(merged[merged["Rank Move"] > 0], ascending=False), make_table(merged[merged["Rank Move"] < 0], ascending=True), previous_label, latest_label


def calculate_ytd_trends(rankings: pd.DataFrame, player_df: pd.DataFrame) -> Tuple[pd.DataFrame, str, str]:
    oldest_sort, latest_sort = rankings["Week Sort"].min(), rankings["Week Sort"].max()
    oldest_label = rankings.loc[rankings["Week Sort"] == oldest_sort, "Week Label"].iloc[0]
    latest_label = rankings.loc[rankings["Week Sort"] == latest_sort, "Week Label"].iloc[0]
    start_df = player_df[player_df["Week Sort"] == oldest_sort].sort_values("Rank")
    end_df = player_df[player_df["Week Sort"] == latest_sort].sort_values("Rank")
    if start_df.empty or end_df.empty:
        missing = []
        if start_df.empty: missing.append(oldest_label)
        if end_df.empty: missing.append(latest_label)
        raise ValueError("Selected player is missing from the workbook comparison week(s): " + ", ".join(missing))
    start, end = start_df.iloc[0], end_df.iloc[0]
    rows = []
    for metric in YTD_TREND_METRICS:
        if metric not in player_df.columns:
            continue
        change = performance_change(metric, start.get(metric), end.get(metric))
        rows.append({"Metric": metric, f"Oldest ({oldest_label})": format_value(metric, start.get(metric)), f"Latest ({latest_label})": format_value(metric, end.get(metric)), "Trend": format_trend(metric, change), "Metric card delta": format_delta_for_metric_card(metric, change)})
    return pd.DataFrame(rows), oldest_label, latest_label


def get_latest_top_n(rankings: pd.DataFrame, n: int) -> Tuple[pd.DataFrame, str]:
    latest_sort = rankings["Week Sort"].max()
    latest_label = rankings.loc[rankings["Week Sort"] == latest_sort, "Week Label"].iloc[0]
    return rankings[rankings["Week Sort"] == latest_sort].copy().sort_values("Rank", na_position="last").head(n), latest_label


def style_distribution_chart(fig):
    fig.update_traces(marker_line_color="rgba(31, 41, 55, 0.95)", marker_line_width=1, texttemplate="%{y}", textposition="outside", cliponaxis=False)
    fig.update_layout(title_x=0.5, uniformtext_minsize=10, uniformtext_mode="show")
    return fig


def chart_metric(player_df: pd.DataFrame, metric: str):
    chart_df = player_df[["Week Label", "Week Sort", metric]].dropna().sort_values("Week Sort")
    if chart_df.empty:
        st.info(f"No numeric data available for {metric} for this player.")
        return
    chart_df = chart_df.copy()
    chart_df["Data Label"] = chart_df[metric].apply(lambda value: format_value(metric, value))
    fig = px.line(chart_df, x="Week Label", y=metric, text="Data Label", markers=True, title=f"{metric} by week", labels={"Week Label": "Week", metric: metric})
    fig.update_traces(textposition="top center", textfont_size=12)
    fig.update_layout(xaxis_title="Week", yaxis_title=metric, hovermode="x unified")
    if metric in LOWER_IS_BETTER:
        fig.update_yaxes(autorange="reversed")
    if "%" in metric:
        fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, width="stretch")


def format_top_players_table(source_df: pd.DataFrame, include_ontario_rank: bool = False) -> pd.DataFrame:
    player_cols = ["Rank", "Player", "Year of Birth", "Points", "Tournaments", "Province", "Club", "Singles WTN", "Career Matches", "Singles W/L-Career", "Singles W/L-YTD", "Singles W/L Career %", "Singles W/L-YTD %"]
    player_cols = [col for col in player_cols if col in source_df.columns]
    sorted_source = source_df.sort_values("Rank", na_position="last").copy()
    if include_ontario_rank:
        sorted_source.insert(sorted_source.columns.get_loc("Rank") + 1 if "Rank" in sorted_source.columns else 0, "Ontario Rank", range(1, len(sorted_source) + 1))
        if "Ontario Rank" not in player_cols:
            player_cols.insert(1 if "Rank" in player_cols else 0, "Ontario Rank")
        if "Province" in player_cols:
            player_cols.remove("Province")
    top_players = sorted_source[player_cols].copy()
    for metric in ["Rank", "Ontario Rank", "Year of Birth", "Points", "Tournaments", "Singles WTN", "Career Matches", "Singles W/L Career %", "Singles W/L-YTD %"]:
        if metric in top_players.columns:
            top_players[metric] = top_players[metric].apply(lambda value, metric=metric: format_value(metric, value))
    return top_players


def render_top_players_table(source_df: pd.DataFrame, title: str, height: int | None, include_ontario_rank: bool = False):
    st.subheader(title, anchor=False)
    display_df = format_top_players_table(source_df, include_ontario_rank=include_ontario_rank)
    st.dataframe(display_df, width="stretch", height=height or table_auto_height(len(display_df)), hide_index=True)


def centered_number_table(df: pd.DataFrame):
    styler = df.style
    if "Number of players" in df.columns:
        styler = styler.set_properties(subset=["Number of players"], **{"text-align": "center"})
    return styler


def render_top_n_analytics(topn_df: pd.DataFrame, latest_label: str, n: int, *, show_province_chart: bool = True, clubs_include_province: bool = True, include_players_table: bool = False, players_table_title: str | None = None, players_table_height: int | None = 760, caption_prefix: str | None = None, include_ontario_rank: bool = False):
    caption = caption_prefix or f"Top {n} analysis"
    st.caption(f"{caption} based on {latest_label} data.")
    yob_counts = topn_df["Year of Birth"].dropna().astype(int).astype(str).value_counts().sort_index().reset_index()
    yob_counts.columns = ["Year of Birth", "Count"]
    club_source = topn_df.copy()
    club_source["Club"] = club_source["Club"].fillna("Unavailable").astype(str).str.strip().replace("", "Unavailable")
    club_source["Province"] = club_source["Province"].fillna("Unavailable").astype(str).str.strip().replace("", "Unavailable") if "Province" in club_source.columns else "Unavailable"
    if clubs_include_province:
        club_counts = (club_source.groupby("Club", as_index=False).agg(Province=("Province", lambda values: ", ".join(sorted(set(values)))), **{"Number of players": ("Player", "nunique")}).sort_values(["Number of players", "Club"], ascending=[False, True]).head(10))
    else:
        club_counts = (club_source.groupby("Club", as_index=False).agg(**{"Number of players": ("Player", "nunique")}).sort_values(["Number of players", "Club"], ascending=[False, True]).head(10))

    if show_province_chart:
        province_counts = topn_df["Province"].fillna("Unavailable").value_counts().reset_index()
        province_counts.columns = ["Province", "Count"]
        pie_col1, pie_col2 = st.columns(2)
        with pie_col1:
            fig_province = px.pie(province_counts, names="Province", values="Count", title=f"Top {n} by Province")
            fig_province.update_layout(title_x=0.5)
            st.plotly_chart(fig_province, width="stretch")
        with pie_col2:
            fig_yob = px.pie(yob_counts, names="Year of Birth", values="Count", title=f"Top {n} by Year of Birth")
            fig_yob.update_layout(title_x=0.5)
            st.plotly_chart(fig_yob, width="stretch")
    else:
        row_col1, row_col2 = st.columns(2)
        with row_col1:
            fig_yob = px.pie(yob_counts, names="Year of Birth", values="Count", title=f"Top {n} by Year of Birth")
            fig_yob.update_layout(title_x=0.5)
            st.plotly_chart(fig_yob, width="stretch")
        with row_col2:
            table_left, table_mid, table_right = st.columns([1, 8, 1])
            with table_mid:
                st.subheader(f"Top 10 clubs by number of Top {n} players", anchor=False)
                st.dataframe(centered_number_table(club_counts), width="content", hide_index=True)
        st.markdown("<div style='height: 2.25rem;'></div>", unsafe_allow_html=True)

    avg_wtn = topn_df["Singles WTN"].dropna().mean() if "Singles WTN" in topn_df.columns else pd.NA
    avg_matches = topn_df["Career Matches"].dropna().mean() if "Career Matches" in topn_df.columns else pd.NA
    avg_ytd = topn_df["Singles W/L-YTD %"].dropna().mean() if "Singles W/L-YTD %" in topn_df.columns else pd.NA
    metric_col1, metric_col2, metric_col3 = st.columns(3)
    metric_col1.metric("Average WTN", "unavailable" if pd.isna(avg_wtn) else f"{avg_wtn:.1f}")
    metric_col2.metric("Average number of career matches", "unavailable" if pd.isna(avg_matches) else f"{avg_matches:.0f}")
    metric_col3.metric("Average Singles W/L YTD", "unavailable" if pd.isna(avg_ytd) else f"{avg_ytd:.0%}")
    hist_col1, hist_col2 = st.columns(2)
    with hist_col1:
        wtn_data = topn_df[["Singles WTN"]].dropna() if "Singles WTN" in topn_df.columns else pd.DataFrame(columns=["Singles WTN"])
        if wtn_data.empty:
            st.info(f"No WTN data available for the Top {n} players.")
        else:
            fig_wtn = px.histogram(wtn_data, x="Singles WTN", title="WTN Distribution", labels={"Singles WTN": "Singles WTN"})
            fig_wtn.update_traces(xbins=dict(size=0.5))
            fig_wtn.update_layout(title_x=0.5, yaxis_title="Number of players")
            st.plotly_chart(style_distribution_chart(fig_wtn), width="stretch")
    with hist_col2:
        match_data = topn_df[["Career Matches"]].dropna() if "Career Matches" in topn_df.columns else pd.DataFrame(columns=["Career Matches"])
        if match_data.empty:
            st.info(f"No career matches data available for the Top {n} players.")
        else:
            fig_matches = px.histogram(match_data, x="Career Matches", title="Career Matches Distribution", labels={"Career Matches": "Career Matches"})
            fig_matches.update_traces(xbins=dict(size=25))
            fig_matches.update_layout(title_x=0.5, yaxis_title="Number of players")
            st.plotly_chart(style_distribution_chart(fig_matches), width="stretch")
    if show_province_chart:
        st.subheader(f"Top 10 clubs by number of Top {n} players", anchor=False)
        st.dataframe(centered_number_table(club_counts), width="content", hide_index=True)
    if include_players_table:
        render_top_players_table(topn_df, players_table_title or f"Top {n} players by rank", players_table_height, include_ontario_rank=include_ontario_rank)


def render_top_n_tab(rankings: pd.DataFrame, n: int):
    topn_df, latest_label = get_latest_top_n(rankings, n)
    render_top_n_analytics(topn_df, latest_label, n, include_players_table=(n == 20), players_table_title="Top 20 players by rank" if n == 20 else None, players_table_height=760)


def render_top_50_ontario_tab(rankings: pd.DataFrame):
    latest_sort = rankings["Week Sort"].max()
    latest_label = rankings.loc[rankings["Week Sort"] == latest_sort, "Week Label"].iloc[0]
    latest_df = rankings[rankings["Week Sort"] == latest_sort].copy()
    ontario_df = latest_df[latest_df["Province"].astype(str).str.strip().eq("Ontario Tennis Association")]
    top50_ontario_df = ontario_df.sort_values("Rank", na_position="last").head(50)
    if top50_ontario_df.empty:
        st.info("No Ontario Tennis Association players were found in the latest week's data.")
        return
    render_top_n_analytics(top50_ontario_df, latest_label, 50, show_province_chart=False, clubs_include_province=False, include_players_table=True, players_table_title="Top 50 Ontario players by rank", players_table_height=None, caption_prefix="Top 50 Ontario analysis", include_ontario_rank=True)


def render_multi_player_download_tab(rankings: pd.DataFrame):
    st.subheader("Multi-Player Raw Data Download", anchor=False)
    st.caption("Paste player names below (each name on a new line) to retrieve all available data across all weeks and download it as a CSV.")

    names_input = st.text_area(
        "Player Names",
        placeholder="Ela Velic\nPlayer Two\nPlayer Three",
        height=150
    )

    if st.button("Retrieve Data", type="primary"):
        if not names_input.strip():
            st.warning("Please enter at least one player name.")
            return

        target_players_raw = [name.strip() for name in names_input.splitlines() if name.strip()]
        
        rankings = rankings.copy()
        rankings["Player_Lower"] = rankings["Player"].str.lower()
        target_players_lower = [name.lower() for name in target_players_raw]

        filtered_df = rankings[rankings["Player_Lower"].isin(target_players_lower)].drop(columns=["Player_Lower"])

        if filtered_df.empty:
            st.error("No data found for any of the specified player names. Please check the spelling and try again.")
            return

        cols_to_drop = ["Profile", "Province", "Club", "Profile URL", "Week Number", "Ranking Year", "Week Sort"]
        filtered_df = filtered_df.drop(columns=[c for c in cols_to_drop if c in filtered_df.columns])

        sort_cols = [c for c in ["Player", "Week Label"] if c in filtered_df.columns]
        if sort_cols:
            filtered_df = filtered_df.sort_values(sort_cols)

        st.success(f"Found {len(filtered_df):,} total records across {filtered_df['Player'].nunique()} player(s).")
        # Display clean dataframe on screen without formula wrappers
        st.dataframe(filtered_df, width="stretch", hide_index=True)

        # Apply Excel text-force formulas only for the CSV download payload
        download_df = filtered_df.copy()
        wl_columns = ["Singles W/L-Career", "Singles W/L-YTD"]
        for col in wl_columns:
            if col in download_df.columns:
                download_df[col] = download_df[col].apply(
                    lambda x: f'="{x}"' if pd.notna(x) and str(x).strip() != "" else x
                )

        csv_data = download_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download Data as CSV",
            data=csv_data,
            file_name="multi_player_raw_data.csv",
            mime="text/csv",
        )


def render_ai_analysis_tab(selected_player: str, player_df: pd.DataFrame, rankings: pd.DataFrame):
    st.subheader(f"AI Analysis for {selected_player}", anchor=False)
    st.caption("Summarizing the player's history, charts, and YTD trends using Google Gemini.")

    if "GEMINI_API_KEY" not in st.secrets:
        st.warning("Google Gemini API key not found. Please add `GEMINI_API_KEY` to your Streamlit secrets.")
        return

    try:
        player_history = build_player_history_display(player_df).to_string()
    except Exception:
        player_history = player_df.to_string()

    try:
        ytd_df, oldest_label, latest_label = calculate_ytd_trends(rankings, player_df)
        ytd_summary = ytd_df.to_string()
    except Exception:
        ytd_summary = "YTD trends data unavailable."

    prompt = f"""
    You are an expert tennis analyst. Analyze the following statistics for the youth tennis player '{selected_player}':
    
    1. Player Match & Ranking History:
    {player_history}
    
    2. Year-to-Date (YTD) Trends:
    {ytd_summary}
    
    Please provide a concise, professional performance summary structured cleanly under these exact markdown headings:
    ### Overall Performance Overview
    ### Key Strengths & Progress
    ### Areas for Improvement
    ### Outlook & Recommendations
    """

    if st.button("Generate AI Analysis", type="primary"):
        with st.spinner("Analyzing player performance with Google Gemini..."):
            try:
                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                response = client.models.generate_content(
                    model="gemini-3.6-flash",
                    contents=prompt,
                )
                st.markdown(response.text)
            except Exception as e:
                st.error(f"Failed to generate AI analysis: {e}")


def render_sidebar_summary(weeks_loaded: int, players_found: int, ranking_rows: int):
    st.markdown("""
        <style>
        section[data-testid="stSidebar"] > div:first-child { padding-bottom: 12rem; }
        .sidebar-summary-fixed { position: fixed; bottom: 0.65rem; left: 0.75rem; width: 250px; max-width: calc(100vw - 1.5rem); padding: 0.65rem 0.75rem; border: none; box-shadow: none; border-radius: 0.5rem; background: var(--secondary-background-color); z-index: 1000; }
        .sidebar-summary-title { font-size: 0.95rem; font-weight: 700; margin-bottom: 0.5rem; }
        .sidebar-summary-row { display: flex; justify-content: space-between; gap: 0.6rem; margin: 0.22rem 0; font-size: 0.82rem; }
        .sidebar-summary-value { font-weight: 700; }
        </style>
        """, unsafe_allow_html=True)
    st.sidebar.markdown(f"""
        <div class="sidebar-summary-fixed">
            <div class="sidebar-summary-title">Database Summary</div>
            <div class="sidebar-summary-row"><span>Number of weeks of Data</span><span class="sidebar-summary-value">{weeks_loaded:,}</span></div>
            <div class="sidebar-summary-row"><span>Total Players</span><span class="sidebar-summary-value">{players_found:,}</span></div>
            <div class="sidebar-summary-row"><span>Total Rows of Data</span><span class="sidebar-summary-value">{ranking_rows:,}</span></div>
        </div>
        """, unsafe_allow_html=True)


def main():
    st.set_page_config(page_title="OTA Girls U14 Ranking Dashboard", layout="wide")
    st.title("OTA Girls U14 Ranking Dashboard")
    if not MASTER_WORKBOOK.exists():
        st.error(f"Could not find {MASTER_WORKBOOK.name}. Place master.xlsx in the same folder as app.py and rerun the app.")
        return
    try:
        rankings = load_rankings(MASTER_WORKBOOK)
    except Exception as exc:
        st.error(f"Could not load workbook: {exc}")
        return
    dataset_line, updated_line = latest_dataset_text(rankings, MASTER_WORKBOOK)
    st.markdown(f"**{dataset_line}**  \n{updated_line}")
    players = sorted(rankings["Player"].dropna().unique())
    default_index = players.index(DEFAULT_PLAYER) if DEFAULT_PLAYER in players else 0
    selected_player = st.sidebar.selectbox("Select player", players, index=default_index)
    render_sidebar_summary(rankings["Week Label"].nunique(), rankings["Player"].nunique(), len(rankings))
    player_df = rankings[rankings["Player"] == selected_player].sort_values("Week Sort")
    profile_url = get_latest_profile_url(player_df)
    
    latest_yob = "unavailable"
    latest_province = "unavailable"
    latest_club = "unavailable"
    latest = player_df.tail(1)
    if not latest.empty:
        latest_row = latest.iloc[0]
        yob_val = latest_row.get("Year of Birth")
        if not pd.isna(yob_val):
            latest_yob = format_value("Year of Birth", yob_val)
        prov_val = latest_row.get("Province")
        if not pd.isna(prov_val) and str(prov_val).strip() != "":
            latest_province = str(prov_val).strip()
        club_val = latest_row.get("Club")
        if not pd.isna(club_val) and str(club_val).strip() != "":
            latest_club = str(club_val).strip()

    if profile_url:
        safe_url = escape(profile_url, quote=True)
        safe_player = escape(selected_player)
        st.markdown(f'<div style="font-size: 1.75rem; font-weight: 600; margin: 0.25rem 0 0.1rem 0; line-height: 1.25;"><a href="{safe_url}" target="_blank" rel="noopener noreferrer" style="text-decoration: underline;">{safe_player}</a></div>', unsafe_allow_html=True)
        st.markdown(f'<div style="font-size: 0.95rem; color: #6b7280; margin-bottom: 0.1rem;">Year of Birth: {latest_yob}</div>', unsafe_allow_html=True)
        st.markdown(f'<div style="font-size: 0.95rem; color: #6b7280; margin-bottom: 0.1rem;">Province: {latest_province}</div>', unsafe_allow_html=True)
        st.markdown(f'<div style="font-size: 0.95rem; color: #6b7280; margin-bottom: 0.75rem;">Club: {latest_club}</div>', unsafe_allow_html=True)
    else:
        st.subheader(selected_player, anchor=False)
        st.markdown(f'<div style="font-size: 0.95rem; color: #6b7280; margin-top: -1rem; margin-bottom: 0.1rem;">Year of Birth: {latest_yob}</div>', unsafe_allow_html=True)
        st.markdown(f'<div style="font-size: 0.95rem; color: #6b7280; margin-bottom: 0.1rem;">Province: {latest_province}</div>', unsafe_allow_html=True)
        st.markdown(f'<div style="font-size: 0.95rem; color: #6b7280; margin-bottom: 0.75rem;">Club: {latest_club}</div>', unsafe_allow_html=True)

    if not latest.empty:
        latest_row = latest.iloc[0]
        provincial_rank = get_latest_provincial_rank(rankings, selected_player)
        k1, k2, k3, k4 = st.columns(4)
        
        canadian_flag_html = '<img src="https://flagcdn.com/w20/ca.png" style="vertical-align: middle; margin-left: 6px;" alt="Canada Flag"/>'
        wtn_logo_html = '<img src="https://worldtennisnumber.com/favicon-16x16.png" style="vertical-align: middle; margin-left: 6px;" alt="WTN Logo"/>'
        
        val_k1 = "unavailable" if pd.isna(latest_row.get("Rank")) else f"{int(latest_row.get('Rank'))}"
        val_k2 = "unavailable" if pd.isna(provincial_rank) else f"{int(provincial_rank)}"
        val_k3 = "unavailable" if pd.isna(latest_row.get("Points")) else f"{latest_row.get('Points'):,.3f}"
        val_k4 = "unavailable" if pd.isna(latest_row.get("Singles WTN")) else f"{latest_row.get('Singles WTN'):.1f}"

        metric_card_css = """
        <div style="background-color: var(--secondary-background-color); padding: 14px 16px; border-radius: 0.5rem; border: 1px solid rgba(128, 128, 128, 0.2);">
            <div style="font-size: 0.85rem; color: var(--text-color); opacity: 0.8; margin-bottom: 4px;">{label}</div>
            <div style="font-size: 1.6rem; font-weight: 600; line-height: 1.2;">{value}</div>
        </div>
        """

        k1.markdown(metric_card_css.format(label=f"Latest rank {canadian_flag_html}", value=val_k1), unsafe_allow_html=True)
        k2.markdown(metric_card_css.format(label="Provincial rank", value=val_k2), unsafe_allow_html=True)
        k3.markdown(metric_card_css.format(label="Latest points", value=val_k3), unsafe_allow_html=True)
        k4.markdown(metric_card_css.format(label=f"Latest WTN {wtn_logo_html}", value=val_k4), unsafe_allow_html=True)
        
    st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)
    tab_table, tab_charts, tab_ytd, tab_top_movers, tab_latest_trends, tab_top_100, tab_top_50, tab_top_20, tab_top_50_ontario, tab_ai, tab_multi_download = st.tabs([
        "Player history", "Player Charts", "YTD Player trends", "Top movers", "Latest Week Trends", "Top 100", "Top 50", "Top 20", "Top 50 Ontario", "AI Analysis", "Multi-Player Download"
    ])
    with tab_table:
        st.caption("Most recent week appears first. Numeric metric cells include week-to-week trend arrows. For Rank and Singles WTN, lower values are treated as better.")
        player_history_display = build_player_history_display(player_df)
        st.dataframe(style_trend_table(player_history_display, NUMERIC_TREND_COLUMNS), width="stretch", hide_index=True)
    with tab_charts:
        available_metrics = [m for m in OPTIONAL_CHART_METRICS if m in rankings.columns]
        selected_metrics = st.multiselect("Metrics to chart", available_metrics, default=[m for m in DEFAULT_CHART_METRICS if m in available_metrics])
        for metric in selected_metrics:
            chart_metric(player_df, metric)
    with tab_ytd:
        st.markdown("Compares the selected player's numeric metrics between the oldest and most recent weekly sheets in the workbook. For Rank and Singles WTN, lower values are treated as better.")
        try:
            trend_df, oldest_label, latest_label = calculate_ytd_trends(rankings, player_df)
            st.caption(f"Comparison: {oldest_label} → {latest_label}")
            card_metrics = ["Rank", "Points", "Tournaments", "Singles WTN", "Career Matches", "Singles W/L-YTD %"]
            card_df = trend_df[trend_df["Metric"].isin(card_metrics)]
            for i in range(0, len(card_df), 3):
                cols = st.columns(3)
                for col, (_, row) in zip(cols, card_df.iloc[i : i + 3].iterrows()):
                    col.metric(row["Metric"], row[f"Latest ({latest_label})"], delta=row["Metric card delta"], delta_color="normal")
            st.dataframe(trend_df.drop(columns=["Metric card delta"]), width="content", hide_index=True)
        except ValueError as exc:
            st.warning(str(exc))
            st.info("Tip: this tab needs the selected player to exist in both the oldest and most recent weekly sheets.")
    with tab_top_movers:
        try:
            moved_up, moved_down, previous_label, latest_label = build_top_movers(rankings)
            st.caption(f"Ranking movement comparison: {previous_label} → {latest_label}.")
            st.subheader("Top 10 ranking movers up", anchor=False)
            st.dataframe(style_trend_table(moved_up, ["Rank Trend"] + TOP_MOVER_METRIC_COLUMNS), width="stretch", hide_index=True)
            st.subheader("Top 10 ranking movers down", anchor=False)
            st.dataframe(style_trend_table(moved_down, ["Rank Trend"] + TOP_MOVER_METRIC_COLUMNS), width="stretch", hide_index=True)
        except ValueError as exc:
            st.warning(str(exc))
    with tab_latest_trends:
        try:
            latest_trends_df, previous_label, latest_label = build_latest_week_trends(rankings)
            st.caption(f"Comparison: {previous_label} → {latest_label}. Only players listed in {latest_label} are included.")
            st.dataframe(style_trend_table(latest_trends_df, LATEST_TRENDS_COLUMNS), width="stretch", height=520, hide_index=True)
        except ValueError as exc:
            st.warning(str(exc))
    with tab_top_100:
        render_top_n_tab(rankings, 100)
    with tab_top_50:
        render_top_n_tab(rankings, 50)
    with tab_top_20:
        render_top_n_tab(rankings, 20)
    with tab_top_50_ontario:
        render_top_50_ontario_tab(rankings)
    with tab_ai:
        render_ai_analysis_tab(selected_player, player_df, rankings)
    with tab_multi_download:
        render_multi_player_download_tab(rankings)


if __name__ == "__main__":
    main()
