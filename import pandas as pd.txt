import pandas as pd
import streamlit as st
from io import BytesIO
from rapidfuzz import process  # optional, only if you want fuzzy matching

def normalize(name: str) -> str:
    """
    Normalize guest names: strip whitespace, lowercase,
    remove punctuation/diacritics so matching is case-insensitive.
    """
    s = str(name).strip().lower()
    # keep letters, numbers, and spaces only
    return "".join(ch for ch in s if ch.isalnum() or ch.isspace())

# ──────────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────────
st.set_page_config(page_title="Night‐Stay Reconciliation Bot", layout="wide")
st.title("📊 Night‐Stay Reconciliation Bot")

st.markdown(
    """
    Upload your **System.xlsx** and **Booking.com.xlsx** files below.
    The app will:
    1. Normalize guest names (case‐insensitive).
    2. Count nights per guest in each file.
    3. Show a full comparison + highlight mismatches.
    4. Let you download an Excel report with two sheets.
    """
)

col1, col2 = st.columns(2)
with col1:
    sys_file = st.file_uploader("Upload System.xlsx", type=["xlsx"])
with col2:
    bcom_file = st.file_uploader("Upload Booking.com.xlsx", type=["xlsx"])

# ──────────────────────────────────────────────────
# Processing
# ──────────────────────────────────────────────────
if sys_file and bcom_file:
    # Read the first sheet of each workbook
    df_sys = pd.read_excel(sys_file)
    df_b   = pd.read_excel(bcom_file)

    # Normalize names
    df_sys["_GUEST"] = df_sys["Guest Name"].apply(normalize)
    df_b  ["_GUEST"] = df_b["Guest Name"].apply(normalize)

    # Count nights from System: one row = one night
    sys_counts = (
        df_sys
        .groupby("_GUEST")
        .size()
        .reset_index(name="System Nights")
    )

    # Sum nights from Booking.com: use the “Room Night” column
    # (or adjust to your actual “Total of Nights” column name)
    bcom_counts = (
        df_b
        .groupby("_GUEST")["Room Night"]
        .sum()
        .reset_index(name="Booking Nights")
    )

    # Merge, fill missing with 0, compute difference & status
    merged = (
        pd.merge(sys_counts, bcom_counts, on="_GUEST", how="outer")
          .fillna(0)
    )
    merged["System Nights"]  = merged["System Nights"].astype(int)
    merged["Booking Nights"] = merged["Booking Nights"].astype(int)
    merged["Δ Nights"]       = merged["System Nights"] - merged["Booking Nights"]
    merged["Status"]         = merged["Δ Nights"].apply(lambda x: "Match" if x == 0 else "Mismatch")

    # Restore a human-readable guest name column (title‐cased)
    merged.insert(
        loc=0,
        column="Guest",
        value=merged["_GUEST"].str.upper()  # or .str.title() if you prefer Title Case
    )
    merged.drop(columns=["_GUEST"], inplace=True)

    # Display
    st.subheader("🔍 Full Comparison")
    st.dataframe(merged.style.format({"Δ Nights": "{:+d}"}), height=400)

    st.subheader("❗ Mismatches Only")
    df_mismatch = merged.query("Status == 'Mismatch'")
    st.dataframe(df_mismatch, height=300)

    # Prepare download
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Master Comparison", index=False)
        df_mismatch.to_excel(writer, sheet_name="Mismatches", index=False)
    buffer.seek(0)

    st.download_button(
        label="📥 Download Reconciliation Report",
        data=buffer,
        file_name="reconciliation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
