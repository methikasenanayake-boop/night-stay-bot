import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Night-Stay Reconciliation", layout="wide")
st.title("📊 Night-Stay Reconciliation (One Row per Guest)")

# 1. File uploads
sys_file = st.file_uploader("Upload System Excel (.xlsx)", type="xlsx")
bkg_file = st.file_uploader("Upload Booking.com Excel (.xlsx)", type="xlsx")

if sys_file and bkg_file:
    # 2. Read Excel files
    df_sys = pd.read_excel(sys_file, engine="openpyxl")
    df_bkg = pd.read_excel(bkg_file, engine="openpyxl")

    # 3. Identify key columns by position
    sys_name_col    = df_sys.columns[2]   # 3rd column = System guest name
    bkg_name_col    = df_bkg.columns[3]   # 4th column = Booking.com guest name
    bkg_date_col    = df_bkg.columns[4]   # 5th column = Check-in date
    bkg_nights_col  = df_bkg.columns[8]   # 9th column = Booking nights

    # 4. Normalize guest names for matching
    normalize = lambda x: str(x).strip().upper()
    df_sys["_GUEST"] = df_sys[sys_name_col].apply(normalize)
    df_bkg["_GUEST"] = df_bkg[bkg_name_col].apply(normalize)

    # 5. Compute earliest check-in date per guest
    df_bkg[bkg_date_col] = pd.to_datetime(df_bkg[bkg_date_col], errors="coerce")
    arrival = (
        df_bkg
        .groupby("_GUEST")[bkg_date_col]
        .min()
        .dt.date
        .reset_index(name="Check-in Date")
    )

    # 6. Sum Booking.com nights per guest
    df_bkg[bkg_nights_col] = pd.to_numeric(df_bkg[bkg_nights_col], errors="coerce").fillna(0)
    bkg_nights = (
        df_bkg
        .groupby("_GUEST")[bkg_nights_col]
        .sum()
        .reset_index(name="Booking Night")
    )

    # 7. Count System nights per guest
    sys_nights = (
        df_sys
        .groupby("_GUEST")
        .size()
        .reset_index(name="System Night")
    )

    # 8. Merge all metrics into one report
    report = (
        sys_nights
        .merge(arrival,    on="_GUEST", how="outer")
        .merge(bkg_nights, on="_GUEST", how="outer")
        .fillna(0)
    )
    report["System Night"]  = report["System Night"].astype(int)
    report["Booking Night"] = report["Booking Night"].astype(int)

    # 9. Calculate net difference and status
    report["Net Night Difference"] = (
        report["System Night"] - report["Booking Night"]
    )
    report["Status"] = report["Net Night Difference"].apply(
        lambda d: "Match"         if d == 0
                  else "System Extra" if d > 0
                  else "Booking Extra"
    )

    # 10. Final formatting & display
    report = report.rename(columns={"_GUEST": "Guest Name"})
    report = report[
        [
            "Guest Name",
            "Check-in Date",
            "System Night",
            "Booking Night",
            "Net Night Difference",
            "Status",
        ]
    ]

    st.subheader("Full Report")
    st.dataframe(report, height=400)

    # 11. Excel download helper
    def to_excel(df: pd.DataFrame) -> bytes:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Report")
        buffer.seek(0)
        return buffer.getvalue()

    # 12. Download button
    st.download_button(
        "📥 Download Report",
        data=to_excel(report),
        file_name="reconciliation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
