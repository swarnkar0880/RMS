from collections import Counter
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Station Demurrage Reason Summariser",
    page_icon="🚂",
    layout="wide",
)

REASON_CATEGORIES = {
    "Wet & Sticky Coal / Muddy Coal": [
        "wet", "sticky", "muddy", "slush", "rain", "monsoon", "moisture",
    ],
    "Heavy Boulders / Foreign Material": [
        "boulder", "metal", "iron", "stone", "oversize", "foreign", "choke", "jam",
    ],
    "Bunching & Placement Delay": [
        "bunching", "placement delay", "delayed placement", "occupied",
        "cascading", "waiting", "line clear",
    ],
    "Wagon Tippler (WT) Issue": [
        "wt-", "wt ", "tippler", "hydraulic", "power pack", "clamp", "tippling",
    ],
    "CHP / Conveyor / Crusher Issue": [
        "chp", "belt", "conveyor", "crusher", "hopper", "feeder", "stream",
        "ptw", "bunker", "trip",
    ],
    "MGR / Loco / Track / Wagon Defect": [
        "mgr", "loco", "track", "wagon", "detach", "attach", "coupling",
        "derail", "brake", "sick",
    ],
    "Data / Entry Error": [
        "wrong entry", "delete", "duplicate",
    ],
}


def categorize_remark(text: str) -> list[str]:
    if not text:
        return []
    text_lower = text.lower()
    matched = [
        category
        for category, keywords in REASON_CATEGORIES.items()
        if any(kw in text_lower for kw in keywords)
    ]
    return matched if matched else ["Other / Miscellaneous"]


def generate_executive_summary(
    station_df: pd.DataFrame, station_name: str, category_counts: Counter
) -> str:
    total_rakes = len(station_df)
    rakes_with_remarks = int(station_df["has_remark"].sum())

    has_debit = "debit_hours" in station_df.columns
    if has_debit:
        debit_series = pd.to_numeric(station_df["debit_hours"], errors="coerce").fillna(0)
        demurrage_rakes = int((debit_series > 0).sum())
        total_debit_hrs = float(debit_series.sum())
    else:
        demurrage_rakes = 0
        total_debit_hrs = 0.0

    if rakes_with_remarks == 0:
        return (
            f"**{station_name}** has **{total_rakes}** recorded rakes, with no remarks "
            f"logged in Column AC (`MGR_Remarks`) or Column AD (`CHP_Remark`)."
        )

    top_reasons = category_counts.most_common(3)
    reasons_text = ", ".join(
        f"**{cat}** ({cnt} mentions)" for cat, cnt in top_reasons
    )

    summary = (
        f"For **{station_name}**, **{rakes_with_remarks}** of **{total_rakes}** rakes "
        f"have remarks in Column AC (`MGR_Remarks`) and/or Column AD (`CHP_Remark`). "
    )
    if has_debit:
        summary += (
            f"**{demurrage_rakes}** rakes recorded demurrage/debit hours "
            f"(**{total_debit_hrs:,.2f} hrs** total). "
        )
    summary += f"Top abstracted delay reasons: {reasons_text}."
    return summary


@st.cache_data
def load_and_clean_data(file_bytes: bytes, file_name: str, sheet_name: str | None = None):
    from io import BytesIO

    buf = BytesIO(file_bytes)
    xls = pd.ExcelFile(buf)
    target_sheet = sheet_name if sheet_name in xls.sheet_names else xls.sheet_names[0]

    buf.seek(0)
    df = pd.read_excel(buf, sheet_name=target_sheet)

    if df.shape[1] < 30:
        raise ValueError(
            f"Expected at least 30 columns (through Column AD), found {df.shape[1]}."
        )

    col_a_name = str(df.columns[0])
    col_ac_name = str(df.columns[28])
    col_ad_name = str(df.columns[29])

    df["Station_Clean"] = df.iloc[:, 0].astype(str).str.strip()
    df = df[
        df["Station_Clean"].notna()
        & (df["Station_Clean"] != "")
        & (df["Station_Clean"].str.lower() != "nan")
    ].copy()

    df["AC_Clean"] = (
        df.iloc[:, 28]
        .fillna("")
        .astype(str)
        .apply(lambda x: re.sub(r"\s+", " ", x).strip())
    )
    df["AD_Clean"] = (
        df.iloc[:, 29]
        .fillna("")
        .astype(str)
        .apply(lambda x: re.sub(r"\s+", " ", x).strip())
    )

    def combine_remarks(row) -> str:
        parts = []
        if row["AC_Clean"]:
            parts.append(f"[MGR] {row['AC_Clean']}")
        if row["AD_Clean"]:
            parts.append(f"[CHP] {row['AD_Clean']}")
        return " | ".join(parts)

    df["Combined_Remarks"] = df.apply(combine_remarks, axis=1)
    df["has_remark"] = (df["AC_Clean"] != "") | (df["AD_Clean"] != "")
    df["Categories"] = (
        (df["AC_Clean"] + " " + df["AD_Clean"]).str.strip().apply(categorize_remark)
    )

    return df, xls.sheet_names, target_sheet, (col_a_name, col_ac_name, col_ad_name)


st.title("🚂 Station Demurrage & Delay Reason Abstractor")
st.markdown(
    "Upload an Excel file to abstract remarks from **Column AC (`MGR_Remarks`)** and "
    "**Column AD (`CHP_Remark`)** and summarise delay reasons by **Column A (`station`)**."
)

uploaded_file = st.file_uploader("Upload Excel File (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        from io import BytesIO

        raw_bytes = uploaded_file.getvalue()
        sheet_list = pd.ExcelFile(BytesIO(raw_bytes)).sheet_names
        selected_sheet = st.sidebar.selectbox("Select Worksheet", sheet_list, index=0)

        df, _, active_sheet, col_names = load_and_clean_data(
            raw_bytes, uploaded_file.name, selected_sheet
        )

        st.sidebar.caption(
            f"Active sheet: `{active_sheet}`\n\n"
            f"- **Col A:** `{col_names[0]}`\n"
            f"- **Col AC:** `{col_names[1]}`\n"
            f"- **Col AD:** `{col_names[2]}`"
        )

        stations = sorted(df["Station_Clean"].unique())
        selected_station = st.sidebar.selectbox("Select Station (Column A)", stations)

        only_demurrage = False
        if "debit_hours" in df.columns:
            only_demurrage = st.sidebar.checkbox(
                "Filter only rakes with Demurrage (debit_hours > 0)", value=False
            )

        only_with_remarks = st.sidebar.checkbox(
            "Show only rows with non-empty AC/AD remarks", value=True
        )

        station_df = df[df["Station_Clean"] == selected_station].copy()

        category_counts = Counter()
        for cats in station_df["Categories"]:
            category_counts.update(cats)

        st.subheader(f"📋 Demurrage & Delay Summary — {selected_station}")
        st.info(generate_executive_summary(station_df, selected_station, category_counts))

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Rakes", len(station_df))
        m2.metric("Rakes with AC Remarks", int((station_df["AC_Clean"] != "").sum()))
        m3.metric("Rakes with AD Remarks", int((station_df["AD_Clean"] != "").sum()))
        if "debit_hours" in station_df.columns:
            total_debit = float(
                pd.to_numeric(station_df["debit_hours"], errors="coerce").fillna(0).sum()
            )
            m4.metric("Total Debit Hours", f"{total_debit:,.2f}")
        else:
            m4.metric("Combined Remarked Rakes", int(station_df["has_remark"].sum()))

        left_col, right_col = st.columns([1, 1])

        with left_col:
            st.markdown("#### 📊 Categorised Reason Breakdown")
            if category_counts:
                cat_df = pd.DataFrame(
                    category_counts.most_common(),
                    columns=["Reason Category", "Occurrences"],
                )
                st.bar_chart(cat_df.set_index("Reason Category"), width="stretch")
                st.dataframe(cat_df, width="stretch", hide_index=True)
            else:
                st.write("No categorised remarks found for this station.")

        with right_col:
            st.markdown("#### 🔁 Top Exact Remarks (AC & AD)")
            all_remarks = pd.concat([
                station_df[station_df["AC_Clean"] != ""]["AC_Clean"],
                station_df[station_df["AD_Clean"] != ""]["AD_Clean"],
            ])
            if not all_remarks.empty:
                top_exact = all_remarks.str.upper().value_counts().reset_index()
                top_exact.columns = ["Exact Remark Text", "Frequency"]
                st.dataframe(top_exact.head(15), width="stretch", hide_index=True)
            else:
                st.write("No exact remarks available.")

        st.markdown("---")
        st.subheader("📑 Abstracted Row-Level Data")

        display_df = station_df.copy()
        if only_demurrage and "debit_hours" in display_df.columns:
            display_df = display_df[
                pd.to_numeric(display_df["debit_hours"], errors="coerce").fillna(0) > 0
            ]
        if only_with_remarks:
            display_df = display_df[display_df["has_remark"]]

        context_cols = [
            c
            for c in ["date", "raketype", "source", "noofwagons", "debit_hours"]
            if c in display_df.columns
        ]
        final_cols = ["Station_Clean"] + context_cols + [
            "AC_Clean",
            "AD_Clean",
            "Categories",
        ]

        renamed_display = display_df[final_cols].rename(
            columns={
                "Station_Clean": "Station (Col A)",
                "AC_Clean": f"{col_names[1]} (Col AC)",
                "AD_Clean": f"{col_names[2]} (Col AD)",
                "Categories": "Abstracted Reason Category",
            }
        )
        renamed_display["Abstracted Reason Category"] = renamed_display[
            "Abstracted Reason Category"
        ].apply(lambda x: ", ".join(x))

        st.dataframe(renamed_display, width="stretch", hide_index=True)

        csv_data = renamed_display.to_csv(index=False).encode("utf-8")
        safe_station = re.sub(r"[^A-Za-z0-9_-]+", "_", selected_station).strip("_")
        st.download_button(
            label=f"⬇️ Download Abstracted Report for {selected_station} (CSV)",
            data=csv_data,
            file_name=f"{safe_station}_demurrage_summary.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")
