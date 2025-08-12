import io
import os

import streamlit as st
import pandas as pd
import openai

# ————————————————
# Setup
# ————————————————
st.set_page_config(page_title="Night‐Stay Reconciliation + Copilot", layout="wide")
st.title("📊 Night‐Stay Reconciliation Bot + Copilot Assistant")

# Load OpenAI key from Secrets
openai.api_key = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY"))

# ————————————————
# 1) File Upload
# ————————————————
sys_file = st.file_uploader("Upload System Excel (.xlsx)", type="xlsx")
bkg_file = st.file_uploader("Upload Booking.com Excel (.xlsx)", type="xlsx")

if not (sys_file and bkg_file):
    st.info("Please upload BOTH System and Booking.com files to proceed.")
    st.stop()

# ————————————————
# 2) Read & Aggregate
# ————————————————
df_sys = pd.read_excel(sys_file, engine="openpyxl")
df_bkg = pd.read_excel(bkg_file, engine="openpyxl")

# auto‐select by index
sys_name, bkg_name = df_sys.columns[2], df_bkg.columns[3]
bkg_arr, bkg_nights = df_bkg.columns[4], df_bkg.columns[8]

# normalize names
norm = lambda x: str(x).strip().upper()
df_sys["_GUEST"] = df_sys[sys_name].apply(norm)
df_bkg["_GUEST"] = df_bkg[bkg_name].apply(norm)

# parse and aggregate
df_bkg[bkg_arr]     = pd.to_datetime(df_bkg[bkg_arr], errors="coerce")
df_bkg[bkg_nights]  = pd.to_numeric(df_bkg[bkg_nights], errors="coerce").fillna(0)

sys_agg = df_sys.groupby("_GUEST").size().rename("System Nights").reset_index()
bkg_agg = (
    df_bkg
    .groupby("_GUEST")
    .agg(**{
        "Booking Nights": (bkg_nights, "sum"),
        "Arrival Date":   (bkg_arr, "min")
    })
    .reset_index()
)
report = (
    sys_agg.merge(bkg_agg, on="_GUEST", how="outer")
    .fillna({"System Nights": 0, "Booking Nights": 0})
)
report["System Nights"]  = report["System Nights"].astype(int)
report["Booking Nights"] = report["Booking Nights"].astype(int)
report["Arrival Date"]   = report["Arrival Date"].dt.date
report["Δ Nights"]       = report["Booking Nights"] - report["System Nights"]
report["Status"]         = report["Δ Nights"].map(lambda d: "Match" if d==0 else ("Booking > System" if d>0 else "System > Booking"))
report = report.rename(columns={"_GUEST":"Guest"})
report = report[["Guest","Arrival Date","System Nights","Booking Nights","Δ Nights","Status"]]

# ————————————————
# 3) Display Reports
# ————————————————
col1, col2 = st.columns(2)
with col1:
    st.subheader("📋 Full Report")
    st.dataframe(report, use_container_width=True)

    # download helper
    def to_excel(df):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Report")
        buf.seek(0)
        return buf.getvalue()

    st.download_button("📥 Download Report", to_excel(report),
                       "reconciliation_report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# mismatches & overlaps
mism = report[report["Status"]!="Match"]
ovlp = report[(report["System Nights"]>0)&(report["Booking Nights"]>0)]
with col2:
    st.subheader("❗ Mismatch Report")
    st.dataframe(mism, use_container_width=True)
    st.subheader("🔄 Overlap Report")
    st.dataframe(ovlp, use_container_width=True)

# ————————————————
# 4) Copilot Chat Panel
# ————————————————
st.sidebar.markdown("## 🤖 Copilot Assistant")
user_query = st.sidebar.text_input("Ask Copilot about the data…", "")

if user_query:
    # build a minimal context: column names + top 5 rows
    context = (
        "Columns: " + ", ".join(report.columns.tolist()) + "\n"
        "Top 5 rows:\n" + report.head().to_csv(index=False)
    )
    prompt = (
        "You are a data assistant. "
        "I have a reconciliation report with these columns and sample rows:\n\n"
        f"{context}\n\n"
        f"User question: {user_query}\n\n"
        "Please answer concisely and refer back to the data where needed."
    )
    # call OpenAI
    try:
        resp = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role":"user","content":prompt}],
            temperature=0.2,
            max_tokens=300
        )
        answer = resp.choices[0].message.content.strip()
    except Exception as e:
        answer = f"Error calling Copilot: {e}"

    st.sidebar.markdown("### Copilot’s Answer")
    st.sidebar.write(answer)
