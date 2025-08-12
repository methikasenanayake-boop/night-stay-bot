import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Night‐Stay Reconciliation", layout="wide")
st.title("📊 Night‐Stay Reconciliation Bot")

st.markdown(
    """
    Upload your two Excel files (.xlsx).  
    Then use the sidebar to map your actual column names.
    """
)

# 1) File uploads
sys_file = st.file_uploader("System file (.xlsx)", type=["xlsx"], key="sys")
bkg_file = st.file_uploader("Booking.com file (.xlsx)", type=["xlsx"], key="bkg")

if sys_file and bkg_file:
    # 2) Read raw sheets
    try:
        df_sys_raw = pd.read_excel(sys_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Could not read System file: {e}")
        st.stop()

    try:
        df_bkg_raw = pd.read_excel(bkg_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Could not read Booking.com file: {e}")
        st.stop()

    # 3) Column‐mapping in sidebar
    st.sidebar.header("Column Mapping")

    guest_sys_col = st.sidebar.selectbox(
        "System: Guest Name column",
        df_sys_raw.columns.tolist()
    )

    guest_bkg_col = st.sidebar.selectbox(
        "Booking: Guest Name column",
        df_bkg_raw.columns.tolist()
    )

    date_bkg_col = st.sidebar.selectbox(
        "Booking: Check-in Date column",
        df_bkg_raw.columns.tolist()
    )

    nights_bkg_col = st.sidebar.selectbox(
        "Booking: Room Nights column",
        df_bkg_raw.columns.tolist()
    )

    # 4) Normalize function
    def normalize(val):
        return str(val).strip().upper()

    # 5) Apply mappings
    df_sys = df_sys_raw.copy()
    df_bkg = df_bkg_raw.copy()

    df_sys["_GUEST"] = df_sys[guest_sys_col].apply(normalize)
    df_bkg["_GUEST"] = df_bkg[guest_bkg_col].apply(normalize)

    # 6) Arrival dates
    df_bkg[date_bkg_col] = pd.to_datetime(df_bkg[date_bkg_col], errors="coerce")
    arrival = (
        df_bkg
        .groupby("_GUEST")[date_bkg_col]
        .min()
        .dt.date
        .reset_index(name="Arrival Date")
    )

    # 7) Night counts
    sys_cnt = df_sys.groupby("_GUEST").size().reset_index(name="System Nights")
    bkg_cnt = df_bkg.groupby("_GUEST")[nights_bkg_col].sum().reset_index(name="Booking Nights")

    # 8) Merge & compute
    df = sys_cnt.merge(bkg_cnt, on="_GUEST", how="outer")
    df = df.merge(arrival, on="_GUEST", how="left").fillna(0)
    df[["System Nights","Booking Nights"]] = df[["System Nights","Booking Nights"]].astype(int)
    df["Δ Nights"] = df["Booking Nights"] - df["System Nights"]
    df["Status"] = df["Δ Nights"].apply(
        lambda x: "Match" if x == 0 else ("System Missing" if x > 0 else "System Extra")
    )

    # 9) Finalize
    df = df.rename(columns={"_GUEST":"Guest"})
    df = df[["Guest","Arrival Date","System Nights","Booking Nights","Δ Nights","Status"]]

    # 10) Split
    full       = df
    mismatches = df[df["Status"] != "Match"]
    overlaps   = df[(df["System Nights"] > 0) & (df["Booking Nights"] > 0)]

    # 11) Display
    st.subheader("🔍 Full Report")
    st.dataframe(full, height=300)

    st.subheader("❗ Mismatch Report")
    st.dataframe(mismatches, height=200)

    st.subheader("🔄 Overlap Report")
    st.dataframe(overlaps, height=200)

    # 12) Download helper
    def to_excel(df_):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_.to_excel(writer, index=False, sheet_name="Report")
        return buf.getvalue()

    st.download_button("📥 Download Full Report",
                       data=to_excel(full),
                       file_name="full_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.download_button("📥 Download Mismatch Report",
                       data=to_excel(mismatches),
                       file_name="mismatch_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet")

    st.download_button("📥 Download Overlap Report",
                       data=to_excel(overlaps),
                       file_name="overlap_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet")
