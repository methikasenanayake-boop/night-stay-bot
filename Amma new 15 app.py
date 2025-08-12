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

    # 3. Identify needed columns by index
    sys_name_col   = df_sys.columns[2]  # 3rd column = guest name in system
    bkg_guest_col  = df_bkg.columns[3]  # 4th column = guest name in booking
    bkg_date_col   = df_bkg.columns[4]  # 5th column = check-in date
    bkg_nights_col = df_bkg.columns[8]  # 9th column = total nights

    # 4. Normalize guest names for reliable joins
    normalize = lambda x: str(x).strip().upper()
    df_sys["_GUEST"] = df_sys[sys_name_col].apply(normalize)
    df_bkg["_GUEST"] = df_bkg[bkg_guest_col].apply(normalize)

    # 5. Extract earliest check-in date per guest
    df_bkg[bkg_date_col] = pd.to_datetime(df_bkg[bkg_date_col], errors="coerce")
    arrival = (
        df_bkg
        .groupby("_GUEST")[bkg_date_col]
        .min()
        .dt.date
        .reset_index(name="Checking Date")
    )

    # 6. Count system nights per guest
    sys_cnt = (
        df_sys
        .groupby("_GUEST")
        .size()
        .reset_index(name="System Night")
    )

    # 7. Sum booking nights per guest
    df_bkg[bkg_nights_col] = pd.to_numeric(df_bkg[bkg_nights_col], errors="coerce").fillna(0)
    bkg_cnt = (
        df_bkg
        .groupby("_GUEST")[bkg_nights_col]
        .sum()
        .reset_index(name="Booking Night")
    )

    # 8. Merge aggregates into one report
    report = (
        sys_cnt
        .merge(arrival, on="_GUEST", how="left")
        .merge(bkg_cnt, on="_GUEST", how="left")
        .fillna({"Booking Night": 0})
    )

    # 9. Calculate differences and status
    report["System Night"]  = report["System Night"].astype(int)
    report["Booking Night"] = report["Booking Night"].astype(int)
    report["Net Night Difference"] = (
        report["System Night"] - report["Booking Night"]
    )
    report["Status"] = report["Net Night Difference"].apply(
        lambda d: "Match" if d == 0 else ("System Extra" if d > 0 else "Booking Extra")
    )

    # 10. Rename and reorder for final display
    report = report.rename(columns={"_GUEST": "Guest Name"})
    report = report[
        [
            "Guest Name",
            "Checking Date",
            "System Night",
            "Booking Night",
            "Net Night Difference",
            "Status"
        ]
    ]

    # 11. Display in Streamlit and offer download
    st.subheader("📋 Reconciliation Report")
    st.dataframe(report, height=400)

    def to_excel(df: pd.DataFrame) -> bytes:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Report")
        buffer.seek(0)
        return buffer.getvalue()

    st.download_button(
        label="📥 Download Report",
        data=to_excel(report),
        file_name="night_stay_reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
