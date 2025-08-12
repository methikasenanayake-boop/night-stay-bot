import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Night‐Stay Reconciliation", layout="wide")
st.title("📊 Night‐Stay Reconciliation (One Row per Guest)")

# 1. File uploads
sys_file = st.file_uploader("Upload System Excel (.xlsx)", type="xlsx")
bkg_file = st.file_uploader("Upload Booking.com Excel (.xlsx)", type="xlsx")

if sys_file and bkg_file:
    # 2. Read Excel files
    df_sys = pd.read_excel(sys_file, engine="openpyxl")
    df_bkg = pd.read_excel(bkg_file, engine="openpyxl")

    # 3. Identify needed columns by index
    sys_name_col   = df_sys.columns[2]   # 3rd col = guest name
    bkg_date_col   = df_bkg.columns[4]   # 5th col = check-in date
    bkg_nights_col = df_bkg.columns[8]   # 9th col = total nights

    # 4. Normalize guest names
    norm = lambda x: str(x).strip().upper()
    df_sys["_GUEST"] = df_sys[sys_name_col].apply(norm)
    df_bkg["_GUEST"] = df_bkg[df_bkg.columns[3]].apply(norm)

    # 5. Earliest check-in date per guest
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

    # 8. Merge all aggregates
    report = (
        sys_cnt
        .merge(arrival, on="_GUEST", how="left")
        .merge(bkg_cnt, on="_GUEST", how="left")
        .fillna({"Booking Night": 0})
    )

    # 9. Calculate net difference and status in one line
    report["System Night"]  = report["System Night"].astype(int)
    report["Booking Night"] = report["Booking Night"].astype(int)
    report["Net Night Difference"] = report["System Night"] - report["Booking Night"]
    report["Status"] = report["Net Night Difference"].apply(lambda d: "Match" if d == 0 else ("System Extra" if d > 0 else "Booking Extra"))

    # 10. Rename and reorder columns
    report = report.rename(columns={"