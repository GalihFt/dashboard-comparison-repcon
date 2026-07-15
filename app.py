from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import streamlit as st


SOURCE_FILE = Path("Pre-Post Jan-Mei 2026.xlsx")
PREPARED_SUMMARY_XLSX = Path("temp/summary_pre_post.xlsx")
PREPARED_DETAIL_XLSX = Path("temp/detail_comparison_pre_post.xlsx")
DATA_DIR = Path("data")
DEPLOY_SUMMARY = DATA_DIR / "summary.parquet"
DEPLOY_DETAIL = DATA_DIR / "detail.parquet"
CACHE_DIR = Path("temp/streamlit_cache")
SUMMARY_CACHE = CACHE_DIR / "summary.parquet"
DETAIL_CACHE = CACHE_DIR / "detail.parquet"


SUMMARY_KEYS = ["NOCONTAINER", "NO_EOR"]
DETAIL_KEYS = [
    "NO_EOR",
    "NOCONTAINER",
    "MATERIAL",
    "COMPONENT",
    "LOCATION",
    "DAMAGE",
    "REPAIRACTION",
]

DETAIL_DISPLAY_FIELDS = [
    "JENISEOR",
    "CLAIM",
    "PAIDBY",
    "MATERIAL",
    "COMPONENT",
    "LOCATION",
    "DAMAGE",
    "DAMAGE_BY",
    "REPAIRACTION",
    "SIZEMATERIAL",
    "QTY",
    "TOTALMHRACTUAL",
    "TOTALLABOURCOSTACTUAL",
    "TOTALCOSTMATERIALACTUAL",
    "SUBTOTALACTUAL",
]

CLAIM_COMPARISON_COLUMNS = [
    *SUMMARY_KEYS,
    "PRE_CLAIM",
    "POST_CLAIM",
    "PRE_SUBTOTALACTUAL",
    "POST_SUBTOTALACTUAL",
]


st.set_page_config(
    page_title="Pre vs Post Survey Monitor",
    page_icon="",
    layout="wide",
)


st.markdown(
    """
    <style>
    .block-container { padding-top: 1.25rem; }
    [data-testid="stMetricValue"] { font-size: 1.35rem; }
    div[data-testid="stDataFrame"] { border: 1px solid #d8dee9; border-radius: 8px; }
    </style>
    """,
    unsafe_allow_html=True,
)


def parse_timestamp(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def first_notna(series: pd.Series):
    values = series.dropna()
    return values.iloc[0] if len(values) else pd.NA


def clean_cost_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in ["TOTALLABOURCOSTACTUAL", "TOTALCOSTMATERIALACTUAL"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def prepare_summary_from_raw(data: pd.DataFrame) -> pd.DataFrame:
    pre = data[data["SURVEYTYPE"] == "PRE SURVEY"].copy()
    post = data[data["SURVEYTYPE"] == "POST SURVEY"].copy()

    def summarize_side(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
        df = clean_cost_columns(df)
        return (
            df.groupby(SUMMARY_KEYS, dropna=False)
            .agg(
                **{
                    f"{prefix}_BANYAK_PEKERJAAN": ("NO_EOR", "size"),
                    f"{prefix}_TOTALLABOURCOSTACTUAL": (
                        "TOTALLABOURCOSTACTUAL",
                        "sum",
                    ),
                    f"{prefix}_TOTALCOSTMATERIALACTUAL": (
                        "TOTALCOSTMATERIALACTUAL",
                        "sum",
                    ),
                    f"{prefix}_TSTAMP": ("TSTAMP", first_notna),
                }
            )
            .reset_index()
        )

    meta_cols = [col for col in ["TSTAMP", "VESSELVOYAGE", "PORTID"] if col in data.columns]
    meta = data.groupby(SUMMARY_KEYS, dropna=False).agg({col: first_notna for col in meta_cols}).reset_index()
    summary = (
        meta.merge(summarize_side(pre, "PRE"), on=SUMMARY_KEYS, how="outer")
        .merge(summarize_side(post, "POST"), on=SUMMARY_KEYS, how="outer")
    )

    number_cols = [col for col in summary.columns if col.startswith(("PRE_", "POST_"))]
    number_cols = [col for col in number_cols if col not in ["PRE_TSTAMP", "POST_TSTAMP"]]
    summary[number_cols] = summary[number_cols].fillna(0)
    summary["SELISIH"] = (
        summary["POST_TOTALLABOURCOSTACTUAL"]
        + summary["POST_TOTALCOSTMATERIALACTUAL"]
        - summary["PRE_TOTALLABOURCOSTACTUAL"]
        - summary["PRE_TOTALCOSTMATERIALACTUAL"]
    )
    return summary


def prepare_detail_from_raw(data: pd.DataFrame) -> pd.DataFrame:
    pre = data[data["SURVEYTYPE"] == "PRE SURVEY"].copy()
    post = data[data["SURVEYTYPE"] == "POST SURVEY"].copy()
    compare_cols = [
        "TSTAMP",
        "VESSELVOYAGE",
        "PORTID",
        "DATEIN",
        "JENISEOR",
        "CLAIM",
        "PAIDBY",
        "DAMAGE_BY",
        "SIZEMATERIAL",
        "QTY",
        "MHRSPIL",
        "MHRVENDOR",
        "LABOURCOSTACTUAL",
        "TOTALMHRACTUAL",
        "TOTALLABOURCOSTACTUAL",
        "COSTMATERIALACTUAL",
        "TOTALCOSTMATERIALACTUAL",
        "SUBTOTALACTUAL",
        "BYUSER",
        "APPTYPE",
    ]
    compare_cols = [col for col in compare_cols if col in data.columns]

    def prepare_side(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
        side = df[DETAIL_KEYS + compare_cols].copy()
        for col in DETAIL_KEYS:
            if side[col].dtype == "object":
                side[col] = side[col].astype("string").str.strip()
        side["MATCH_NO"] = side.groupby(DETAIL_KEYS, dropna=False).cumcount() + 1
        return side.rename(columns={col: f"{prefix}_{col}" for col in compare_cols})

    detail = prepare_side(pre, "PRE").merge(
        prepare_side(post, "POST"),
        on=DETAIL_KEYS + ["MATCH_NO"],
        how="outer",
        indicator=True,
    )
    detail["MATCH_STATUS"] = detail["_merge"].map(
        {"both": "MATCH", "left_only": "PRE_ONLY", "right_only": "POST_ONLY"}
    )
    return detail.drop(columns=["_merge"])


def read_source_workbook() -> pd.DataFrame:
    sheets = pd.read_excel(SOURCE_FILE, sheet_name=None)
    return pd.concat(sheets.values(), ignore_index=True)


def normalize_loaded_frames(summary: pd.DataFrame, detail: pd.DataFrame):
    for col in ["TSTAMP", "PRE_TSTAMP", "POST_TSTAMP"]:
        if col in summary.columns:
            summary[col] = parse_timestamp(summary[col])
        if col in detail.columns:
            detail[col] = parse_timestamp(detail[col])

    if "TSTAMP" not in summary.columns:
        summary["TSTAMP"] = summary.get("PRE_TSTAMP")
    summary["FILTER_TSTAMP"] = summary["TSTAMP"]
    if "PRE_TSTAMP" in summary.columns:
        summary["FILTER_TSTAMP"] = summary["FILTER_TSTAMP"].fillna(summary["PRE_TSTAMP"])
    if "POST_TSTAMP" in summary.columns:
        summary["FILTER_TSTAMP"] = summary["FILTER_TSTAMP"].fillna(summary["POST_TSTAMP"])

    number_cols = [
        "PRE_BANYAK_PEKERJAAN",
        "PRE_TOTALLABOURCOSTACTUAL",
        "PRE_TOTALCOSTMATERIALACTUAL",
        "POST_BANYAK_PEKERJAAN",
        "POST_TOTALLABOURCOSTACTUAL",
        "POST_TOTALCOSTMATERIALACTUAL",
        "SELISIH",
    ]
    for col in number_cols:
        if col in summary.columns:
            summary[col] = pd.to_numeric(summary[col], errors="coerce").fillna(0)

    return summary, detail


def read_parquet_compact(path: Path) -> pd.DataFrame:
    return pd.read_parquet(path, dtype_backend="pyarrow")


def save_cache(summary: pd.DataFrame, detail: pd.DataFrame) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    summary.to_parquet(SUMMARY_CACHE, index=False)
    detail.to_parquet(DETAIL_CACHE, index=False)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    summary.to_parquet(DEPLOY_SUMMARY, index=False)
    detail.to_parquet(DEPLOY_DETAIL, index=False)


def has_required_detail_columns(detail: pd.DataFrame) -> bool:
    required_base = ["JENISEOR", "CLAIM", "DAMAGE_BY", "SIZEMATERIAL"]
    required = {f"PRE_{col}" for col in required_base}
    required |= {f"POST_{col}" for col in required_base}
    return required.issubset(detail.columns)


def build_cache() -> tuple[pd.DataFrame, pd.DataFrame]:
    if PREPARED_SUMMARY_XLSX.exists() and PREPARED_DETAIL_XLSX.exists():
        summary = pd.read_excel(PREPARED_SUMMARY_XLSX)
        detail = pd.read_excel(PREPARED_DETAIL_XLSX)
        if not has_required_detail_columns(detail):
            data = read_source_workbook()
            summary = prepare_summary_from_raw(data)
            detail = prepare_detail_from_raw(data)
    else:
        data = read_source_workbook()
        summary = prepare_summary_from_raw(data)
        detail = prepare_detail_from_raw(data)

    summary, detail = normalize_loaded_frames(summary, detail)
    save_cache(summary, detail)
    return summary, detail


@st.cache_resource(show_spinner=False)
def load_data(force_rebuild: bool = False) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not force_rebuild and DEPLOY_SUMMARY.exists() and DEPLOY_DETAIL.exists():
        summary = read_parquet_compact(DEPLOY_SUMMARY)
        detail = read_parquet_compact(DEPLOY_DETAIL)
        if has_required_detail_columns(detail):
            return normalize_loaded_frames(summary, detail)

    if not force_rebuild and SUMMARY_CACHE.exists() and DETAIL_CACHE.exists():
        summary = read_parquet_compact(SUMMARY_CACHE)
        detail = read_parquet_compact(DETAIL_CACHE)
        if not has_required_detail_columns(detail):
            return build_cache()
        return normalize_loaded_frames(summary, detail)
    return build_cache()


def format_currency(value) -> str:
    try:
        return f"{float(value):,.0f}"
    except (TypeError, ValueError):
        return "0"


def format_signed_number(value) -> str:
    try:
        return f"{float(value):+,.0f}"
    except (TypeError, ValueError):
        return "+0"


def filter_summary_by_date(summary: pd.DataFrame, date_range) -> pd.DataFrame:
    if not date_range or len(date_range) != 2:
        return summary
    start_date, end_date = date_range
    start_ts = pd.Timestamp(start_date)
    end_ts = pd.Timestamp(end_date) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
    return summary[
        summary["FILTER_TSTAMP"].isna()
        | summary["FILTER_TSTAMP"].between(start_ts, end_ts)
    ].copy()


def parse_no_eor_queries(value: str) -> list[str]:
    queries = [query.strip() for query in re.split(r"[\n,;]+", value) if query.strip()]
    unique_queries = {}
    for query in queries:
        unique_queries.setdefault(query.casefold(), query)
    return list(unique_queries.values())


def filter_detail_by_no_eor(detail: pd.DataFrame, value: str) -> pd.DataFrame:
    queries = parse_no_eor_queries(value)
    if not queries:
        return detail.iloc[0:0].copy()

    normalized_no_eor = detail["NO_EOR"].astype("string").str.casefold()
    matches = pd.Series(False, index=detail.index)
    for query in queries:
        matches |= normalized_no_eor.str.contains(query.casefold(), na=False, regex=False)
    return detail[matches].copy()


def build_claim_comparison(
    detail: pd.DataFrame,
    filtered_summary: pd.DataFrame,
) -> pd.DataFrame:
    active_keys = filtered_summary[SUMMARY_KEYS].drop_duplicates()
    active_detail = detail[CLAIM_COMPARISON_COLUMNS].merge(
        active_keys,
        on=SUMMARY_KEYS,
        how="inner",
    )
    rows = []

    for claim_type in ["CLAIM", "NON CLAIM"]:
        for survey_type in ["PRE", "POST"]:
            claim_column = f"{survey_type}_CLAIM"
            subtotal_column = f"{survey_type}_SUBTOTALACTUAL"
            normalized_claim = (
                active_detail[claim_column].astype("string").str.strip().str.upper()
            )
            claim_detail = active_detail[normalized_claim == claim_type]
            rows.append(
                {
                    "CLAIM_TYPE": claim_type,
                    "SURVEY_TYPE": survey_type,
                    "EOR_CONTAINER": len(claim_detail[SUMMARY_KEYS].drop_duplicates()),
                    "PEKERJAAN": len(claim_detail),
                    "TOTAL_BIAYA": pd.to_numeric(
                        claim_detail[subtotal_column], errors="coerce"
                    ).sum(),
                }
            )

    return pd.DataFrame(rows).set_index(["CLAIM_TYPE", "SURVEY_TYPE"])


def metric_delta(post_value, pre_value, formatter=format_signed_number) -> str:
    return f"{formatter(post_value - pre_value)} vs PRE"


def format_detail_display(detail: pd.DataFrame) -> pd.DataFrame:
    display = pd.DataFrame()
    empty = pd.Series(pd.NA, index=detail.index)
    zero = pd.Series(0, index=detail.index)
    pre_exists = detail["MATCH_STATUS"] != "POST_ONLY"
    post_exists = detail["MATCH_STATUS"] != "PRE_ONLY"

    display["NOCONTAINER"] = detail["NOCONTAINER"]
    display["NO_EOR"] = detail["NO_EOR"]
    display["TSTAMP"] = detail.get("PRE_TSTAMP", empty).combine_first(
        detail.get("POST_TSTAMP", empty)
    )
    display["VESSELVOYAGE"] = detail.get(
        "PRE_VESSELVOYAGE", empty
    ).combine_first(detail.get("POST_VESSELVOYAGE", empty))
    display["PORTID"] = detail.get("PRE_PORTID", empty).combine_first(
        detail.get("POST_PORTID", empty)
    )

    for prefix in ["PRE", "POST"]:
        side_exists = pre_exists if prefix == "PRE" else post_exists
        for col in DETAIL_DISPLAY_FIELDS:
            source_col = f"{prefix}_{col}"
            if source_col in detail.columns:
                display[source_col] = detail[source_col]
            elif col in DETAIL_KEYS:
                display[source_col] = detail[col].where(side_exists)

    pre_subtotal = pd.to_numeric(
        detail.get("PRE_SUBTOTALACTUAL", zero),
        errors="coerce",
    ).fillna(0)
    post_subtotal = pd.to_numeric(
        detail.get("POST_SUBTOTALACTUAL", zero),
        errors="coerce",
    ).fillna(0)
    display["SELISIH"] = post_subtotal - pre_subtotal

    return display


with st.sidebar:
    st.header("Filter")
    can_rebuild = (
        SOURCE_FILE.exists()
        or (PREPARED_SUMMARY_XLSX.exists() and PREPARED_DETAIL_XLSX.exists())
    )
    rebuild = st.button(
        "Refresh cache dari data",
        use_container_width=True,
        disabled=not can_rebuild,
    )

summary_df, detail_df = load_data(force_rebuild=rebuild)

valid_dates = summary_df["FILTER_TSTAMP"].dropna()
if valid_dates.empty:
    default_range = None
else:
    default_range = (valid_dates.min().date(), valid_dates.max().date())

with st.sidebar:
    selected_dates = st.date_input(
        "Range TSTAMP",
        value=default_range,
        min_value=default_range[0] if default_range else None,
        max_value=default_range[1] if default_range else None,
    )
    only_large = st.number_input(
        "Minimal selisih absolut",
        min_value=0,
        value=0,
        step=100000,
    )
    exclude_zero_post = st.checkbox(
        "Exclude POST pekerjaan = 0",
        value=True,
    )
    with st.form("detail_search_form"):
        detail_no_eor_input = st.text_area(
            "Search detail by NO_EOR",
            placeholder=(
                "Satu atau beberapa NO_EOR\n"
                "Contoh: EOR/00003012/01/2026, EOR/00003013/01/2026"
            ),
            height=100,
        )
        search_submitted = st.form_submit_button("Search", use_container_width=True)
    if search_submitted:
        st.session_state["detail_no_eor_query"] = detail_no_eor_input.strip()
    detail_no_eor_search = st.session_state.get("detail_no_eor_query", "")
    status_filter = st.multiselect(
        "Status detail",
        ["MATCH", "PRE_ONLY", "POST_ONLY"],
        default=["MATCH", "PRE_ONLY", "POST_ONLY"],
    )

st.title("Pre vs Post Survey Monitor")

filtered_summary = filter_summary_by_date(summary_df, selected_dates)
if only_large:
    filtered_summary = filtered_summary[filtered_summary["SELISIH"].abs() >= only_large]
if exclude_zero_post:
    filtered_summary = filtered_summary[filtered_summary["POST_BANYAK_PEKERJAAN"] > 0]

filtered_summary = filtered_summary.sort_values("SELISIH", key=lambda s: s.abs(), ascending=False)

pre_total = filtered_summary["PRE_TOTALLABOURCOSTACTUAL"].sum() + filtered_summary[
    "PRE_TOTALCOSTMATERIALACTUAL"
].sum()
post_total = filtered_summary["POST_TOTALLABOURCOSTACTUAL"].sum() + filtered_summary[
    "POST_TOTALCOSTMATERIALACTUAL"
].sum()

metric_cols = st.columns(4)
metric_cols[0].metric("Total EOR/Container", f"{len(filtered_summary):,}")
metric_cols[1].metric("Total Pre", format_currency(pre_total))
metric_cols[2].metric("Total Post", format_currency(post_total))
metric_cols[3].metric("Selisih Post - Pre", format_currency(post_total - pre_total))

claim_comparison = build_claim_comparison(detail_df, filtered_summary)
claim_pre = claim_comparison.loc[("CLAIM", "PRE")]
claim_post = claim_comparison.loc[("CLAIM", "POST")]
non_claim_pre = claim_comparison.loc[("NON CLAIM", "PRE")]
non_claim_post = claim_comparison.loc[("NON CLAIM", "POST")]

st.subheader("Perbandingan Claim dan Non Claim")
st.caption("Mengikuti filter summary yang sedang aktif.")

st.markdown("**PRE Survey**")
pre_volume_cols = st.columns(4)
pre_volume_cols[0].metric(
    "Claim EOR/Container PRE",
    f"{claim_pre['EOR_CONTAINER']:,.0f}",
)
pre_volume_cols[1].metric(
    "Non Claim EOR/Container PRE",
    f"{non_claim_pre['EOR_CONTAINER']:,.0f}",
)
pre_volume_cols[2].metric(
    "Claim Pekerjaan PRE",
    f"{claim_pre['PEKERJAAN']:,.0f}",
)
pre_volume_cols[3].metric(
    "Non Claim Pekerjaan PRE",
    f"{non_claim_pre['PEKERJAAN']:,.0f}",
)

pre_cost_cols = st.columns(2)
pre_cost_cols[0].metric(
    "Claim Total Biaya PRE",
    format_currency(claim_pre["TOTAL_BIAYA"]),
)
pre_cost_cols[1].metric(
    "Non Claim Total Biaya PRE",
    format_currency(non_claim_pre["TOTAL_BIAYA"]),
)

st.markdown("**POST Survey**")
post_volume_cols = st.columns(4)
post_volume_cols[0].metric(
    "Claim EOR/Container POST",
    f"{claim_post['EOR_CONTAINER']:,.0f}",
    metric_delta(claim_post["EOR_CONTAINER"], claim_pre["EOR_CONTAINER"]),
    delta_color="off",
)
post_volume_cols[1].metric(
    "Non Claim EOR/Container POST",
    f"{non_claim_post['EOR_CONTAINER']:,.0f}",
    metric_delta(non_claim_post["EOR_CONTAINER"], non_claim_pre["EOR_CONTAINER"]),
    delta_color="off",
)
post_volume_cols[2].metric(
    "Claim Pekerjaan POST",
    f"{claim_post['PEKERJAAN']:,.0f}",
    metric_delta(claim_post["PEKERJAAN"], claim_pre["PEKERJAAN"]),
    delta_color="off",
)
post_volume_cols[3].metric(
    "Non Claim Pekerjaan POST",
    f"{non_claim_post['PEKERJAAN']:,.0f}",
    metric_delta(non_claim_post["PEKERJAAN"], non_claim_pre["PEKERJAAN"]),
    delta_color="off",
)

post_cost_cols = st.columns(2)
post_cost_cols[0].metric(
    "Claim Total Biaya POST",
    format_currency(claim_post["TOTAL_BIAYA"]),
    metric_delta(claim_post["TOTAL_BIAYA"], claim_pre["TOTAL_BIAYA"]),
    delta_color="off",
)
post_cost_cols[1].metric(
    "Non Claim Total Biaya POST",
    format_currency(non_claim_post["TOTAL_BIAYA"]),
    metric_delta(non_claim_post["TOTAL_BIAYA"], non_claim_pre["TOTAL_BIAYA"]),
    delta_color="off",
)

summary_view_cols = [
    "NO_EOR",
    "NOCONTAINER",
    "TSTAMP",
    "VESSELVOYAGE",
    "PORTID",
    "PRE_BANYAK_PEKERJAAN",
    "PRE_TOTALLABOURCOSTACTUAL",
    "PRE_TOTALCOSTMATERIALACTUAL",
    "POST_BANYAK_PEKERJAAN",
    "POST_TOTALLABOURCOSTACTUAL",
    "POST_TOTALCOSTMATERIALACTUAL",
    "SELISIH",
]
summary_view_cols = [col for col in summary_view_cols if col in filtered_summary.columns]
summary_display = filtered_summary[summary_view_cols].copy()
if "SELISIH" in summary_display.columns:
    summary_display["SELISIH"] = summary_display["SELISIH"].map(format_currency)

st.subheader("Summary per EOR dan Container")
summary_event = st.dataframe(
    summary_display,
    use_container_width=True,
    height=420,
    hide_index=True,
    on_select="rerun",
    selection_mode="single-row",
    column_config={
        "TSTAMP": st.column_config.DatetimeColumn("TSTAMP", format="DD-MM-YYYY HH:mm:ss"),
    },
)

selected_rows = summary_event.selection.rows
if selected_rows:
    selected = filtered_summary.iloc[selected_rows[0]]
else:
    selected = None

st.subheader("Detail one-by-one")
if detail_no_eor_search:
    no_eor_queries = parse_no_eor_queries(detail_no_eor_search)
    selected_detail = filter_detail_by_no_eor(detail_df, detail_no_eor_search)
    st.caption(
        f"Search {len(no_eor_queries)} NO_EOR: "
        + ", ".join(no_eor_queries)
    )
elif selected is None:
    st.info("Klik satu baris summary atau isi NO_EOR lalu tekan Search.")
    selected_detail = pd.DataFrame()
else:
    selected_no_eor = selected["NO_EOR"]
    selected_container = selected["NOCONTAINER"]
    st.caption(f"{selected_no_eor} | {selected_container}")

    selected_detail = detail_df[
        (detail_df["NO_EOR"] == selected_no_eor)
        & (detail_df["NOCONTAINER"] == selected_container)
    ].copy()

if not selected_detail.empty:
    if status_filter:
        selected_detail = selected_detail[selected_detail["MATCH_STATUS"].isin(status_filter)]

    pre_pekerjaan = selected_detail["MATCH_STATUS"].isin(["MATCH", "PRE_ONLY"]).sum()
    post_pekerjaan = selected_detail["MATCH_STATUS"].isin(["MATCH", "POST_ONLY"]).sum()
    selisih_pekerjaan = post_pekerjaan - pre_pekerjaan
    pre_value = pd.to_numeric(
        selected_detail.get("PRE_SUBTOTALACTUAL", pd.Series(dtype="float64")),
        errors="coerce",
    ).sum()
    post_value = pd.to_numeric(
        selected_detail.get("POST_SUBTOTALACTUAL", pd.Series(dtype="float64")),
        errors="coerce",
    ).sum()
    selisih_value = post_value - pre_value

    detail_metrics = st.columns(4)
    detail_metrics[0].metric("Pekerjaan PRE", f"{pre_pekerjaan:,}")
    detail_metrics[1].metric("Pekerjaan POST", f"{post_pekerjaan:,}")
    detail_metrics[2].metric("Selisih pekerjaan", format_signed_number(selisih_pekerjaan))
    detail_metrics[3].metric("Detail rows", f"{len(selected_detail):,}")

    value_metrics = st.columns(3)
    value_metrics[0].metric("Total Biaya PRE", format_currency(pre_value))
    value_metrics[1].metric("Total Biaya POST", format_currency(post_value))
    value_metrics[2].metric("Selisih Biaya", format_signed_number(selisih_value))

    detail_display = format_detail_display(selected_detail)

    st.dataframe(
        detail_display,
        use_container_width=True,
        height=520,
        hide_index=True,
        column_config={
            "TSTAMP": st.column_config.DatetimeColumn("TSTAMP", format="DD-MM-YYYY HH:mm:ss"),
        },
    )
elif detail_no_eor_search:
    st.info("NO_EOR yang dicari tidak ditemukan di detail.")
