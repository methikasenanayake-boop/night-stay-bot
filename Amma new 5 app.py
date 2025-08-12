import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Night-Stay Reconciliation", layout="wide")
st.title("📊 Automated Night-Stay Reconciliation")

# 1. Upload both Excel files
sys_file = st.file_uploader("Upload System Excel (.xlsx)", type="xlsx")
bkg_file = st.file_uploader("Upload Booking.com Excel (.xlsx)", type="xlsx")

if sys_file and bkg_file:
    # 2. Read files with explicit engine
    df_sys = pd.read_excel(sys_file, engine="openpyxl")
    df_bkg = pd.read_excel(bkg_file, engine="openpyxl")

    # 3. Identify columns by index (0-based):
    #    System:  3rd col → guest name
    #    Booking: 4th col → guest name
    #             5th col → arrival date
    #             9th col → total nights
    sys_guest_col   = df_sys.columns[2]
    bkg_guest_col  = df_bkg.columns[3]
    bkg_arrival_col = df_bkg.columns[4]
    bkg_nights_col  = df_bkg.columns[8]

    # 4. Normalize guest names
    def normalize(x):
        return str(x).strip().upper()

    df_sys["_GUEST"] = df_sys[sys_guest_col].apply(normalize)
    df_bkg["_GUEST"] = df_bkg[bkg_guest_col].apply(normalize)

    # 5. Compute earliest arrival date per guest
    df_bkg[bkg_arrival_col] = pd.to_datetime(df_bkg[bkg_arrival_col], errors="coerce")
    arrival = (
        df_bkg
        .groupby("_GUEST")[bkg_arrival_col]
        .min()
        .dt.date
        .reset_index(name="Arrival Date")
    )

    # 6. Count nights
    #    System: one row per night
    sys_cnt = (
        df_sys
        .groupby("_GUEST")
        .size()
        .reset_index(name="System Nights")
    )

    #    Booking.com: sum the 9th column
    df_bkg[bkg_nights_col] = pd.to_numeric(df_bkg[bkg_nights_col], errors="coerce").fillna(0)
    bkg_cnt = (
        df_bkg
        .groupby("_GUEST")[bkg_nights_col]
        .sum()
        .reset_index(name="Booking Nights")
    )

    # 7. Merge, diff & status
    merged = sys_cnt.merge(bkg_cnt, on="_GUEST", how="outer")
    merged = merged.merge(arrival, on="_GUEST", how="left").fillna(0)
    merged[["System Nights","Booking Nights"]] = merged[["System Nights","Booking Nights"]].astype(int)
    merged["Δ Nights"] = merged["Booking Nights"] - merged["System Nights"]
    merged["Status"] = merged["Δ Nights"].apply(
        lambda d: "Match" if d == 0 else ("System Missing" if d > 0 else "System Extra")
    )

    # 8. Finalize columns
    merged = merged.rename(columns={"_GUEST": "Guest"})
    merged = merged[[
        "Guest", "Arrival Date",
        "System Nights", "Booking Nights",
        "Δ Nights", "Status"
    ]]

    # 9. Split reports
    full       = merged
    mismatches = merged[merged["Status"] != "Match"]
    overlaps   = merged[(merged["System Nights"] > 0) & (merged["Booking Nights"] > 0)]

    # 10. Display
    st.subheader("🔍 Full Report")
    st.dataframe(full, height=300)

    st.subheader("❗ Mismatch Report")
    st.dataframe(mismatches, height=200)

    st.subheader("🔄 Overlap Report")
    st.dataframe(overlaps, height=200)

    # 11. Download helper
    def to_excel(df_):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_.to_excel(writer, index=False, sheet_name="Report")
        buf.seek(0)
        return buf.getvalue()

    # 12. Download buttons
    st.download_button(
        "📥 Download Full Report",
        data=to_excel(full),
        file_name="full_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        "📥 Download Mismatch Report",
        data=to_excel(mismatches),
        file_name="mismatch_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet"
    )
    st.download_button(
        "📥 Download Overlap Report",
        data=to_excel(overlaps),
        file_name="overlap_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet"
    )
