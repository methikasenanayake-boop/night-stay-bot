import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Guest Night-Stay Reconciliation", layout="wide")
st.title("Guest Night-Stay Reconciliation Bot")
st.markdown("Upload your System file and Booking.com file to compare guest night-stay records.")

file_system = st.file_uploader("System Excel file", type=["xlsx"])
file_booking = st.file_uploader("Booking.com Excel file", type=["xlsx"])

if file_system and file_booking:
    df_system = pd.read_excel(file_system)
    df_booking = pd.read_excel(file_booking)

    st.subheader("Data Previews")
    st.write("System file", df_system.head())
    st.write("Booking.com file", df_booking.head())

    # Adjust these to match your actual column headers
    key_columns = ["Guest Name", "Month", "Nights"]

    missing = [c for c in key_columns if c not in df_system.columns or c not in df_booking.columns]
    if missing:
        st.error(f"Missing columns in one of the files: {missing}")
        st.stop()

    merged = df_system.merge(
        df_booking,
        on=key_columns,
        how="outer",
        indicator=True
    )

    matched      = merged[merged["_merge"] == "both"].drop(columns=["_merge"])
    only_system  = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])
    only_booking = merged[merged["_merge"] == "right_only"].drop(columns=["_merge"])

    st.subheader("Matched Records")
    st.dataframe(matched)

    st.subheader("Only in System file")
    st.dataframe(only_system)

    st.subheader("Only in Booking.com file")
    st.dataframe(only_booking)

    # Build Excel download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        matched.to_excel(writer, sheet_name="Matched", index=False)
        only_system.to_excel(writer, sheet_name="Only_System", index=False)
        only_booking.to_excel(writer, sheet_name="Only_Booking", index=False)
    data = buffer.getvalue()

    st.download_button(
        label="Download Reconciliation Report",
        data=data,
        file_name="reconciliation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


