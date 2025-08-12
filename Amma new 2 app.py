import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Night‐Stay Reconciliation", layout="wide")
st.title("📊 Night‐Stay Reconciliation Bot")

st.markdown(
    """
    Upload your two Excel files (.xlsx).  
    The app will compare guest night-stay counts, show arrival dates,  
    and produce three downloadable reports: Full, Mismatches, Overlaps.
    """
)

# 1) File upload widgets
sys_file = st.file_uploader("System file (.xlsx)", type=["xlsx"])
bkg_file = st.file_uploader("Booking.com file (.xlsx)", type=["xlsx"])

if sys_file and bkg_file:
    # 2) Read Excel with explicit engine
    try:
        df_sys = pd.read_excel(sys_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Could not read System file: {e}")
        st.stop()

    try:
        df_bkg = pd.read_excel(bkg_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Could not read Booking.com file: {e}")
        st.stop()

    # 3) Normalize guest names
    def normalize(name):
        return str(name).strip().upper()

    df_sys["_GUEST"] = df_sys["Guest Name"].apply(normalize)
    df_bkg["_GUEST"] = df_bkg["Guest Name"].apply(normalize)

    # 4) Compute arrival dates
    df_bkg["Check in Date"] = pd.to_datetime(df_bkg["Check in Date"])
    arrival = (
        df_bkg
        .groupby("_GUEST")["Check in Date"]
        .min()
        .dt.date
        .reset_index(name="Arrival Date")
    )

    # 5) Count nights
    sys_cnt = df_sys.groupby("_GUEST").size().reset_index(name="System Nights")
    bkg_cnt = df_bkg.groupby("_GUEST")["Room Night"].sum().reset_index(name="Booking Nights")

    # 6) Merge and fill
    df = sys_cnt.merge(bkg_cnt, on="_GUEST", how="outer")
    df = df.merge(arrival, on="_GUEST", how="left").fillna(0)
    df[["System Nights","Booking Nights"]] = df[["System Nights","Booking Nights"]].astype(int)

    # 7) Calculate differences & status
    df["Δ Nights"] = df["Booking Nights"] - df["System Nights"]
    df["Status"] = df["Δ Nights"].apply(
        lambda x: "Match" if x == 0 else ("System Missing" if x > 0 else "System Extra")
    )

    # 8) Finalize columns
    df = df.rename(columns={"_GUEST": "Guest"})
    df = df[["Guest","Arrival Date","System Nights","Booking Nights","Δ Nights","Status"]]

    # 9) Split reports
    full       = df
    mismatches = df[df["Status"] != "Match"]
    overlaps   = df[(df["System Nights"] > 0) & (df["Booking Nights"] > 0)]

    # 10) Display tables
    st.subheader("🔍 Full Report")
    st.dataframe(full, height=300)

    st.subheader("❗ Mismatch Report")
    st.dataframe(mismatches, height=200)

    st.subheader("🔄 Overlap Report")
    st.dataframe(overlaps, height=200)

    # 11) Download buttons
    def to_excel(df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Report")
        return buffer.getvalue()

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
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        "📥 Download Overlap Report",
        data=to_excel(overlaps),
        file_name="overlap_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
