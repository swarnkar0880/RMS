import streamlit as st
import pandas as pd
import re
from collections import Counter

# Page configuration
st.set_page_config(
    page_title="Station Demurrage Reason Summariser",
    page_icon="🚂",
    layout="wide"
)

# Domain-specific keyword categories tailored to MGR & CHP remarks
REASON_CATEGORIES = {
    "Wet & Sticky Coal / Muddy Coal": [
        "wet", "sticky", "muddy", "slush", "rain", "monsoon", "moisture"
    ],
    "Heavy Boulders / Foreign Material": [
        "boulder", "metal", "iron", "stone", "oversize", "foreign", "choke", "jam"
    ],
    "Bunching & Placement Delay": [
        "bunching", "placement delay", "delayed placement", "occupied", "cascading", "waiting", "line clear"
    ],
    "Wagon Tippler (WT) Issue": [
        "wt-", "wt ", "tippler", "hydraulic", "power pack", "clamp", "tippling"
    ],
    "CHP / Conveyor / Crusher Issue": [
        "chp", "belt", "conveyor", "crusher", "hopper", "feeder", "stream", "ptw", "bunker", "trip"
    ],
    "MGR / Loco / Track / Wagon Defect": [
        "mgr", "loco", "track", "wagon", "detach", "attach", "coupling", "derail", "brake", "sick"
    ],
    "Data / Entry Error": [
        "wrong entry", "delete", "duplicate"
    ]
}

def categorize_remark(text: str) -> list[str]:
    """Maps a cleaned remark string to one or more operational delay categories."""
    if not text:
        return []
    text_lower = text.lower()
    matched = []
    for category, keywords in REASON_CATEGORIES.items():
        if any(kw in text_lower for kw in keywords):
            matched.append(category)
    return matched if matched else ["Other / Miscellaneous"]

def generate_executive_summary(station_df: pd.DataFrame, station_name: str, category_counts: Counter) -> str:
    """Generates a structured summary paragraph from abstracted AC & AD remarks."""
    total_rakes = len(station_df)
    rakes_with_remarks = station_df["has_remark"].sum()
    
    # Check if debit_hours exists in the uploaded file
    has_debit = "debit_hours" in station_df.columns
    if has_debit:
        station_df["debit_hours"] = pd.to_numeric(station_df["debit_hours"], errors="coerce").fillna(0)
        demurrage_rakes = (station_df["debit_hours"] > 0).sum()
        total_debit_hrs = station_df["debit_hours"].sum()
    else:
        demurrage_rakes = 0
        total_debit_hrs = 0.0

    if rakes_with_remarks == 0:
        return f"**{station_name}** has **{total_rakes}** recorded rakes, but no remarks were entered in Columns AC (`MGR_Remarks`) or AD (`CHP_Remark`)."

    top_reasons = category_counts.most_common(3)
    reasons_text = ", ".join([f"**{cat}** ({cnt} mentions)" for cat, cnt in top_reasons])

    summary = (
        f"For **{station_name}**, out of **{total_rakes}** total rakes, **{rakes_with_remarks}** rakes have "
        f"remarks logged in Columns AC (`MGR_Remarks`) and/or AD (`CHP_Remark`). "
    )
    if has_debit:
        summary += (
            f"A total of **{demurrage_rakes}** rakes incurred demurrage/debit hours (totaling **{total_debit_hrs:,.2f} hrs**). "
        )
    summary += f"The primary abstracted causes of delay are: {reasons_text}."
    return summary

@st.cache_data
def load_and_clean_data(uploaded_file, sheet_name=None):
    xls = pd.ExcelFile(uploaded_file)
    target_sheet = sheet_name if sheet_name else xls.sheet_names[0]
    df = pd.read_excel(uploaded_file, sheet_name=target_sheet)

    # Positional indexing to strictly target Column A (0), Column AC (28), Column AD (29)
    if df.shape[1] < 30:
        raise ValueError(f"Expected at least 30 columns (up to AD), found {df.shape[1]}.")

    col_a_name = df.columns[0]
    col_ac_name = df.columns[28]
    col_ad_name = df.columns[29]

    # Clean whitespace and NaN strings
    df["Station_Clean"] = df.iloc[:, 0].astype(str).str.strip()
    df = df[df["Station_Clean"].notna() & (df["Station_Clean"] != "") & (df["Station_Clean"].str.lower() != "nan")]

    df["AC_Clean"] = df.iloc[:, 28].fillna("").astype(str).apply(lambda x: re.sub(r"\s+", " ", x).strip())
    df["AD_Clean"] = df.iloc[:, 29].fillna("").astype(str).apply(lambda x: re.sub(r"\s+", " ", x).strip())

    # Combine AC and AD into an abstracted remark column
    def combine_remarks(row):
        parts = []
        if row["AC_Clean"]:
            parts.append(f"[MGR] {row['AC_Clean']}")
        if row["AD_Clean"]:
            parts.append(f"[CHP] {row['AD_Clean']}")
        return " | ".join(parts)

    df["Combined_Remarks"] = df.apply(combine_remarks, axis=1)
    df["has_remark"] = (df["AC_Clean"] != "") | (df["AD_Clean"] != "")
    df["Categories"] = (df["AC_Clean"] + " " + df["AD_Clean"]).str.strip().apply(categorize_remark)

    return df, xls.sheet_names, (col_a_name, col_ac_name, col_ad_name)

# --- UI Layout ---
st.title("🚂 Station Demurrage & Delay Reason Abstractor")
st.markdown(
    "Upload your rake unloading Excel sheet to extract remarks from **Column AC (`MGR_Remarks`)** and "
    "**Column AD (`CHP_Remark`)**, and generate a station-wise summary for **Column A (`station`)**."
)

uploaded_file = st.file_uploader("Upload Excel File (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Initial load to get sheet names
        temp_xls = pd.ExcelFile(uploaded_file)
        selected_sheet = st.sidebar.selectbox("Select Worksheet", temp_xls.sheet_names, index=0)

        df, _, col_names = load_and_clean_data(uploaded_file, selected_sheet)
        st.sidebar.caption(f"Mapped Columns:\n- **Col A:** `{col_names[0]}`\n- **Col AC:** `{col_names[1]}`\n- **Col AD:** `{col_names[2]}`")

        # Station selector
        stations = sorted(df["Station_Clean"].unique())
        selected_station = st.sidebar.selectbox("Select Station (Column A)", stations)

        # Optional filter for demurrage-only rakes (debit_hours > 0)
        only_demurrage = False
        if "debit_hours" in df.columns:
            only_demurrage = st.sidebar.checkbox("Filter only rakes with Demurrage (debit_hours > 0)", value=False)

        only_with_remarks = st.sidebar.checkbox("Show only rows with non-empty AC/AD remarks", value=True)

        # Filter dataframe for selected station
        station_df = df[df["Station_Clean"] == selected_station].copy()

        # Count categories across all remarked rows for this station
        category_counts = Counter()
        for cats in station_df["Categories"]:
            category_counts.update(cats)

        # Executive Summary Box
        st.subheader(f"📋 Demurrage & Delay Summary — {selected_station}")
        summary_text = generate_executive_summary(station_df, selected_station, category_counts)
        st.info(summary_text)

        # Key Metrics Row
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Rakes", len(station_df))
        col2.metric("Rakes with AC Remarks", (station_df["AC_Clean"] != "").sum())
        col3.metric("Rakes with AD Remarks", (station_df["AD_Clean"] != "").sum())
        if "debit_hours" in station_df.columns:
            total_debit = pd.to_numeric(station_df["debit_hours"], errors="coerce").fillna(0).sum()
            col4.metric("Total Debit Hours", f"{total_debit:,.2f}")
        else:
            col4.metric("Combined Remarked Rakes", station_df["has_remark"].sum())

        # Breakdown Charts & Tables
        left_col, right_col = st.columns([1, 1])

        with left_col:
            st.markdown("#### 📊 Categorised Reason Breakdown")
            if category_counts:
                cat_df = pd.DataFrame(
                    category_counts.most_common(),
                    columns=["Reason Category", "Occurrences"]
                )
                st.bar_chart(cat_df.set_index("Reason Category"))
                st.dataframe(cat_df, use_container_width=True, hide_index=True)
            else:
                st.write("No categorised remarks found for this station.")

        with right_col:
            st.markdown("#### 🔁 Top Exact Remarks (AC & AD)")
            all_remarks = pd.concat([
                station_df[station_df["AC_Clean"] != ""]["AC_Clean"],
                station_df[station_df["AD_Clean"] != ""]["AD_Clean"]
            ])
            if not all_remarks.empty:
                top_exact = all_remarks.str.upper().value_counts().reset_index()
                top_exact.columns = ["Exact Remark Text", "Frequency"]
                st.dataframe(top_exact.head(15), use_container_width=True, hide_index=True)
            else:
                st.write("No exact remarks available.")

        # Detailed Abstracted Records Table
        st.markdown("---")
        st.subheader("📑 Abstracted Row-Level Data")

        display_df = station_df.copy()
        if only_demurrage and "debit_hours" in display_df.columns:
            display_df = display_df[pd.to_numeric(display_df["debit_hours"], errors="coerce").fillna(0) > 0]
        if only_with_remarks:
            display_df = display_df[display_df["has_remark"]]

        # Select helpful context columns if present
        context_cols = [c for c in ["date", "raketype", "source", "noofwagons", "debit_hours"] if c in display_df.columns]
        final_cols = ["Station_Clean"] + context_cols + ["AC_Clean", "AD_Clean", "Categories"]

        renamed_display = display_df[final_cols].rename(columns={
            "Station_Clean": "Station (Col A)",
            "AC_Clean": f"{col_names[1]} (Col AC)",
            "AD_Clean": f"{col_names[2]} (Col AD)",
            "Categories": "Abstracted Reason Category"
        })
        renamed_display["Abstracted Reason Category"] = renamed_display["Abstracted Reason Category"].apply(lambda x: ", ".join(x))

        st.dataframe(renamed_display, use_container_width=True, hide_index=True)

        # Download CSV button
        csv_data = renamed_display.to_csv(index=False).encode("utf-8")
        st.download_button(
            label=f"⬇️ Download Abstracted Report for {selected_station} (CSV)",
            data=csv_data,
            file_name=f"{selected_station.replace(', ', '_')}_demurrage_summary.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")
