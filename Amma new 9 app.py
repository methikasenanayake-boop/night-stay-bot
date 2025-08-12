import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Night‐Stay Reconciliation", layout="wide")
st.title("📊 Night‐Stay Reconciliation (One Row per Guest)")

# 1. File uploads
sys_file = st.file_uploader("Upload System Excel (.xlsx)", type="xlsx")
bkg_file = st.file_uploader("Upload Booking.com Excel (.xlsx)", type="xlsx")

if sys_file and bkg_file:
    # 2. Read Excel files with openpyxl
    df_sys = pd.read_excel(sys_file, engine="openpyxl")
    df_bkg = pd.read_excel(bkg_file, engine="openpyxl")

    # 3. Column indices (0-based)
    sys_name_col   = df_sys.columns[2]   # 3rd col = guest name
    bkg_guest_col  = df_bkg.columns[3]   # 4th col = guest name
    bkg_date_col   = df_bkg.columns[4]   # 5th col = check‐in date
    bkg_nights_col = df_bkg.columns[8]   # 9th col = booking nights

    # 4. Normalize guest names
    norm = lambda x: str(x).strip().upper()
    df_sys["_GUEST"] = df_sys[sys_name_col].apply(norm)
    df_bkg["_GUEST"] = df_bkg[bkg_guest_col].apply(norm)

    # 5. Compute earliest check-in date per guest
    df_bkg[bkg_date_col] = pd.to_datetime(df_bkg[bkg_date_col], errors="coerce")
    arrival = (
        df_bkg
        .groupby("_GUEST")[bkg_date_col]
        .min()
        .dt.date
        .reset_index(name="Checking Date")
    )

    # 6. Count System nights per guest
    sys_cnt = (
        df_sys
        .groupby("_GUEST")
        .size()
        .reset_index(name="System Night")
    )

    # 7. Sum Booking nights per guest
    df_bkg[bkg_nights_col] = pd.to_numeric(
        df_bkg[bkg_nights_col], errors="coerce"
    ).fillna(0)
    bkg_cnt = (
        df_bkg
        .groupby("_GUEST")[bkg_nights_col]
        .sum()
        .reset_index(name="Booking Night")
    )

    # 8. Merge into one row per guest
    report = (
        sys_cnt
        .merge(arrival, on="_GUEST", how="outer")
        .merge(bkg_cnt, on="_GUEST", how="outer")
        .fillna(0)
    )
    report["System Night"]  = report["System Night"].astype(int)
    report["Booking Night"] = report["Booking Night"].astype(int)

    # 9. Calculate net difference and status
    report["Net Night Difference"] = (
        report["System Night"] - report["Booking Night"]
    )
    report["Status"] = report["Net Night Difference"].apply(
        lambda d: "Match" if d == 0
        else ("System Extra" if d > 0 else "Booking Extra")
    )

    # 10. Rename and reorder columns
    report = report.rename(columns={"_GUEST": "Guest Name"})
    report = report[
        [
            "Guest Name",
            "Checking Date",
            "System Night",
            "Booking Night",
            "Net Night Difference",
            "Status",
        ]
    ]

    # 11. Split into Full / Mismatch / Overlap
    full_report     = report
    mismatch_report = report[report["Net Night Difference"] != 0]
    overlap_report  = report[
        (report["System Night"] > 0) & (report["Booking Night"] > 0)
    ]

    # 12. Display tables
    st.subheader("Full Report")
    st.dataframe(full_report, height=300)

    st.subheader("Mismatch Report")
    st.dataframe(mismatch_report, height=200)

    st.subheader("Overlap Report")
    st.dataframe(overlap_report, height=200)

    # 13. Excel download helper
    def to_excel(df: pd.DataFrame) -> bytes:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Report")
        buffer.seek(0)
        return buffer.getvalue()

    # 14. Download buttons
    st.download_button(
        "📥 Download Full Report",
        data=to_excel(full_report),
        file_name="full_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        "📥 Download Mismatch Report",
        data=to_excel(mismatch_report),
        file_name="mismatch_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        "📥 Download Overlap Report",
        data=to_excel(overlap_report),
        file_name="overlap_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
