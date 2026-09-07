import io
import re
from urllib.parse import urlparse

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Lead List Cleaner",
    page_icon="🧹",
    layout="wide",
)

st.title("Lead List Cleaner")
st.caption("A lightweight business tool for cleaning, standardizing, and quality-checking lead spreadsheets.")

def read_file(uploaded):
    name = uploaded.name.lower()
    data = uploaded.getvalue()
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

    # QA flags before deduplication
    if website_col:
        df["Missing Website"] = df["Website (normalized)"].eq("")
    else:
        df["Missing Website"] = True

    if email_col:
        df["Missing Email"] = df["Email (normalized)"].eq("")
    else:
        df["Missing Email"] = True

    duplicate_key = None
    if company_col and website_col:
        duplicate_key = df["Company (normalized)"] + "|" + df["Website (normalized)"]
    elif website_col:
        duplicate_key = df["Website (normalized)"]
    elif company_col:
        duplicate_key = df["Company (normalized)"]

    if duplicate_key is not None:
        df["Duplicate"] = duplicate_key.duplicated(keep="first")
    else:
        df["Duplicate"] = False

    cleaned = df.loc[~df["Duplicate"]].copy()

    # Remove helper columns from the deliverable while keeping QA flags.
    helper = [c for c in ["Company (normalized)", "Email (normalized)", "Website (normalized)"] if c in cleaned.columns]
    if website_col and website_col not in cleaned.columns:
        cleaned[website_col] = cleaned["Website (normalized)"]
    cleaned = cleaned.drop(columns=helper, errors="ignore")

    return df, cleaned, {
        "company_col": company_col,
        "website_col": website_col,
        "email_col": email_col,
    }

def export_xlsx(cleaned, qa):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="Cleaned Leads", index=False)
        qa.to_excel(writer, sheet_name="Quality Checks", index=False)
    output.seek(0)
    return output

uploaded = st.file_uploader("Upload a lead list", type=["csv", "xlsx"])

if uploaded is None:
    st.info("Upload a CSV or XLSX file to start.")
    st.markdown("### What this demo does")
    st.write("• Standardizes text and URLs  • Flags missing fields  • Detects duplicates  • Produces a cleaned XLSX deliverable")
else:
    try:
        raw = read_file(uploaded)
        raw, cleaned, mapping = clean_dataframe(raw)

        total = len(raw)
        duplicates = int(raw["Duplicate"].sum())
        missing_web = int(raw["Missing Website"].sum())
        missing_email = int(raw["Missing Email"].sum())

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Raw records", total)
        c2.metric("Duplicates flagged", duplicates)
        c3.metric("Missing websites", missing_web)
        c4.metric("Final records", len(cleaned))

        st.markdown("### Data quality")
        qa = pd.DataFrame({
            "Check": [
                "Raw records",
                "Duplicates flagged",
                "Records with missing website",
                "Records with missing email",
                "Final cleaned records",
            ],
            "Result": [total, duplicates, missing_web, missing_email, len(cleaned)],
        })
        st.dataframe(qa, use_container_width=True, hide_index=True)

        st.markdown("### Cleaned output")
        st.dataframe(cleaned, use_container_width=True, hide_index=True)

        export = export_xlsx(cleaned, qa)
        st.download_button(
            "Download cleaned XLSX",
            data=export,
            file_name="cleaned_leads.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.caption(
            f"Detected columns — Company: {mapping['company_col'] or 'not found'} | "
            f"Website: {mapping['website_col'] or 'not found'} | "
            f"Email: {mapping['email_col'] or 'not found'}"
        )

    except Exception as exc:
        st.error(f"Could not process the file: {exc}")
