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
.hero {
    padding: 1.4rem 1.6rem;
    border-radius: 18px;
    background: linear-gradient(135deg,#17365D 0%,#2F75B5 100%);
    color:#fff;
    margin-bottom:1.3rem;
}
.hero h1 {margin:0 0 .35rem 0; font-size:2.3rem;}
.hero p {margin:0; opacity:.92; font-size:1rem;}
.small-note {color:#667085; font-size:.86rem;}
.stButton>button, .stDownloadButton>button {border-radius:10px;}
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
        return pd.read_csv(io.BytesIO(data), sep=None, engine="python")
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

    # Keep the raw input presentation clean: preserve only original columns.
    raw_display = df.copy()

    if website_col:
        norm_url = df[website_col].map(normalize_url)
    else:
        norm_url = pd.Series([""] * len(df), index=df.index)

    if email_col:
        norm_email = df[email_col].str.lower().str.strip()
    else:
        norm_email = pd.Series([""] * len(df), index=df.index)

    if company_col:
        norm_company = df[company_col].str.lower().str.replace(r"[^a-z0-9]+", "", regex=True)
    else:
        norm_company = pd.Series(range(len(df)), index=df.index).astype(str)

    missing_website = norm_url.eq("")
    missing_email = norm_email.eq("") if email_col else pd.Series([False] * len(df), index=df.index)

    if company_col and website_col:
        duplicate_key = norm_company + "|" + norm_url
    elif website_col:
        duplicate_key = norm_url
    elif company_col:
        duplicate_key = norm_company
    else:
        duplicate_key = pd.Series(range(len(df)), index=df.index).astype(str)

    duplicate = duplicate_key.duplicated(keep="first")

    # Build a business-facing cleaned output.
    keep_mask = ~duplicate

    # Reset indexes before assigning derived columns.
    # This avoids pandas alignment errors when an input file contains
    # duplicate row labels or duplicate records.
    cleaned = df.loc[keep_mask].copy().reset_index(drop=True)

    if website_col:
        cleaned[website_col] = norm_url.loc[keep_mask].reset_index(drop=True).to_numpy()
    if email_col:
        cleaned[email_col] = norm_email.loc[keep_mask].reset_index(drop=True).to_numpy()

    # QA summary only — no internal helper columns exposed to the client.
    qa = {
        "total": len(df),
        "duplicates": int(duplicate.sum()),
        "missing_website": int(missing_website.sum()),
        "missing_email": int(missing_email.sum()),
        "final_records": len(cleaned),
        "has_email_column": email_col is not None,
    }

    review_mask = duplicate | missing_website
    if email_col is not None:
        review_mask = review_mask | missing_email
    raw_display["QA Status"] = "OK"
    raw_display.loc[review_mask & ~duplicate, "QA Status"] = "Review"
    raw_display.loc[duplicate, "QA Status"] = "Duplicate"

    return raw_display, cleaned, qa

def export_xlsx(cleaned, qa):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="Cleaned Leads", index=False)
        pd.DataFrame({
            "Check": ["Raw records", "Duplicates flagged", "Records with missing website",
                      "Records with missing email" if qa["has_email_column"] else "Email check",
                      "Final cleaned records"],
            "Result": [qa["total"], qa["duplicates"], qa["missing_website"],
                       qa["missing_email"] if qa["has_email_column"] else "Not provided",
                       qa["final_records"]],
            "Status": ["PASS", "PASS", "REVIEW" if qa["missing_website"] else "PASS",
                       "REVIEW" if qa["has_email_column"] and qa["missing_email"] else "PASS",
                       "PASS"],
        }).to_excel(writer, sheet_name="Quality Checks", index=False)
    output.seek(0)
    return output.getvalue()

sample_path = Path(__file__).with_name("demo_leads.csv")

left, right = st.columns([2, 1])
with left:
    st.subheader("Upload your lead list")
    uploaded = st.file_uploader("CSV or XLSX", type=["csv", "xlsx"], label_visibility="collapsed")
with right:
    st.subheader("Quick demo")
    st.markdown(
        '<div class="small-note">A demo dataset is loaded automatically so visitors can see the complete workflow immediately.</div>',
        unsafe_allow_html=True,
    )
    if sample_path.exists():
        st.download_button(
            "Download sample CSV",
            sample_path.read_bytes(),
            file_name="demo_leads.csv",
            mime="text/csv",
            use_container_width=True,
        )

if uploaded is not None:
    source_name = uploaded.name
    try:
        input_df = read_uploaded(uploaded)
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")
        st.stop()
else:
    source_name = "Demo dataset"
    if sample_path.exists():
        input_df = pd.read_csv(sample_path)
    else:
        st.info("Upload a CSV or XLSX file to start.")
        st.stop()

raw, cleaned, qa = clean_dataframe(input_df)

st.caption(f"Showing: {source_name}")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Raw records", qa["total"])
c2.metric("Duplicates flagged", qa["duplicates"])
c3.metric("Needs review", qa["missing_website"] + (qa["missing_email"] if qa["has_email_column"] else 0))
c4.metric("Final records", qa["final_records"])

st.divider()

tab1, tab2, tab3 = st.tabs(["Quality check", "Cleaned output", "How it works"])

with tab1:
    st.subheader("Quality check")
    checks = [
        ("Raw records", qa["total"], "PASS"),
        ("Duplicates flagged", qa["duplicates"], "PASS"),
        ("Records with missing website", qa["missing_website"], "REVIEW" if qa["missing_website"] else "PASS"),
    ]
    if qa["has_email_column"]:
        checks.append(("Records with missing email", qa["missing_email"], "REVIEW" if qa["missing_email"] else "PASS"))
    else:
        checks.append(("Email field", "Not provided", "PASS"))
    checks.append(("Final cleaned records", qa["final_records"], "PASS"))

    qa_table = pd.DataFrame(checks, columns=["Check", "Result", "Status"])
    st.dataframe(qa_table, use_container_width=True, hide_index=True)

    st.subheader("Raw input")
    st.caption("QA status is added for clarity; internal normalization fields are hidden.")
    st.dataframe(raw, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Cleaned output")
    st.dataframe(cleaned, use_container_width=True, hide_index=True)
    export = export_xlsx(cleaned, qa)
    st.download_button(
        "Download cleaned XLSX",
        export,
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
