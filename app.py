import io
import re
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Lead List Cleaner", page_icon="🧹", layout="wide")

st.markdown("""
<style>
.block-container {max-width: 1200px; padding-top: 2rem; padding-bottom: 3rem;}
.hero {padding: 1.4rem 1.6rem; border-radius: 18px; background: linear-gradient(135deg,#17365D 0%,#2F75B5 100%); color:#fff; margin-bottom:1.2rem;}
.hero h1 {margin:0 0 .35rem 0; font-size:2.3rem;}
.hero p {margin:0; opacity:.92; font-size:1rem;}
.small-note {color:#667085; font-size:.86rem;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>Lead List Cleaner</h1>
<p>Clean, validate and standardize lead spreadsheets — then export a ready-to-use XLSX file.</p>
</div>
""", unsafe_allow_html=True)

def read_uploaded(file_obj):
    data = file_obj.getvalue()
    name = file_obj.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(data))
    if name.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(data))
    raise ValueError("Please upload a CSV or XLSX file.")

def normalize_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()

def normalize_url(value):
    value = normalize_text(value)
    if not value:
        return ""
    if not re.match(r"^https?://", value, flags=re.I):
        value = "https://" + value
    parsed = urlparse(value)
    if not parsed.netloc:
        return ""
    host = parsed.netloc.lower().removeprefix("www.")
    path = parsed.path.rstrip("/")
    return f"https://{host}{path}"

def find_col(df, candidates):
    normalized = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        if candidate in normalized:
            return normalized[candidate]
    for key, original in normalized.items():
        for candidate in candidates:
            if candidate in key:
                return original
    return None

def clean_dataframe(df):
    df = df.copy()
    df.columns = [normalize_text(c) for c in df.columns]
    for col in df.columns:
        df[col] = df[col].map(normalize_text)

    company_col = find_col(df, ["company", "company name", "business", "organization"])
    website_col = find_col(df, ["website", "website url", "url", "domain"])
    email_col = find_col(df, ["email", "email address", "e-mail"])

    if website_col:
        df["Website (normalized)"] = df[website_col].map(normalize_url)
    if email_col:
        df["Email (normalized)"] = df[email_col].str.lower().str.strip()
    if company_col:
        df["Company (normalized)"] = df[company_col].str.lower().str.replace(r"[^a-z0-9]+", "", regex=True)

    df["Missing Website"] = df["Website (normalized)"].eq("") if website_col else True
    df["Missing Email"] = df["Email (normalized)"].eq("") if email_col else True

    if company_col and website_col:
        duplicate_key = df["Company (normalized)"] + "|" + df["Website (normalized)"]
    elif website_col:
        duplicate_key = df["Website (normalized)"]
    elif company_col:
        duplicate_key = df["Company (normalized)"]
    else:
        duplicate_key = pd.Series(range(len(df)), index=df.index).astype(str)

    df["Duplicate"] = duplicate_key.duplicated(keep="first")
    cleaned = df.loc[~df["Duplicate"]].copy()
    cleaned = cleaned.drop(columns=["Company (normalized)", "Email (normalized)", "Website (normalized)"], errors="ignore")
    return df, cleaned

def export_xlsx(cleaned, qa):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="Cleaned Leads", index=False)
        qa.to_excel(writer, sheet_name="Quality Checks", index=False)
    output.seek(0)
    return output

sample_path = Path(__file__).with_name("demo_leads.csv")

left, right = st.columns([2, 1])
with left:
    st.subheader("Upload your lead list")
    uploaded = st.file_uploader("CSV or XLSX", type=["csv", "xlsx"], label_visibility="collapsed")
with right:
    st.subheader("Quick demo")
    if sample_path.exists():
        st.download_button(
            "Download sample input",
            sample_path.read_bytes(),
            file_name="demo_leads.csv",
            mime="text/csv",
            use_container_width=True,
        )
    st.markdown(
        '<div class="small-note">A demo dataset is loaded automatically, so visitors can see the workflow immediately.</div>',
        unsafe_allow_html=True,
    )

if uploaded is not None:
    source_name = uploaded.name
    try:
        input_df = read_uploaded(uploaded)
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")
        st.stop()
elif sample_path.exists():
    source_name = "Demo dataset"
    input_df = pd.read_csv(sample_path)
else:
    st.info("Upload a CSV or XLSX file to start.")
    st.stop()

raw, cleaned = clean_dataframe(input_df)
total = len(raw)
duplicates = int(raw["Duplicate"].sum())
missing_web = int(raw["Missing Website"].sum())
missing_email = int(raw["Missing Email"].sum())

st.caption(f"Showing: {source_name}")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Raw records", total)
c2.metric("Duplicates flagged", duplicates)
c3.metric("Missing websites", missing_web)
c4.metric("Final records", len(cleaned))

tab1, tab2, tab3 = st.tabs(["Quality check", "Cleaned output", "How it works"])

qa = pd.DataFrame({
    "Check": ["Raw records", "Duplicates flagged", "Records with missing website", "Records with missing email", "Final cleaned records"],
    "Result": [total, duplicates, missing_web, missing_email, len(cleaned)],
    "Status": ["PASS", "PASS", "REVIEW" if missing_web else "PASS", "REVIEW" if missing_email else "PASS", "PASS"],
})

with tab1:
    st.subheader("Quality check")
    st.dataframe(qa, use_container_width=True, hide_index=True)
    st.subheader("Raw input")
    st.dataframe(raw, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Cleaned output")
    st.dataframe(cleaned, use_container_width=True, hide_index=True)
    export = export_xlsx(cleaned, qa)
    st.download_button(
        "Download cleaned XLSX",
        export.getvalue(),
        file_name="cleaned_leads.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

with tab3:
    st.subheader("Workflow")
    st.markdown(
        "**1. Collect** — Start with a CSV/XLSX lead list.\n\n"
        "**2. Standardize** — Normalize text and website URLs.\n\n"
        "**3. Validate** — Flag missing fields and potential duplicates.\n\n"
        "**4. Clean** — Remove duplicate records and prepare the final dataset.\n\n"
        "**5. Export** — Download a business-ready XLSX deliverable."
    )
    st.caption("Portfolio demo: self-created sample using public company information. No client relationship is claimed.")

st.divider()
st.caption("Python • Streamlit • Excel/XLSX • Data Cleaning • Quality Checks")
