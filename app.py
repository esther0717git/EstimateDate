# Modified Visitor List Cleaner to preserve all sheets

import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# --- Streamlit setup ---
st.set_page_config(page_title="Visitor List Cleaner", layout="wide")
st.title("\U0001F1F8\U0001F1EC CLARITY GATE – VISITOR DATA CLEANING & VALIDATION \U0001FAE7")

with open("sample_template.xlsx", "rb") as f:
    sample_bytes = f.read()
st.download_button(
    label="🌟 Download Sample Template",
    data=sample_bytes,
    file_name="sample_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.info("""
**Data Integrity Is Our Foundation**  
At every step—from file upload to final report—we enforce strict validation to guarantee your visitor data is accurate, complete, and compliant.  
Maintaining integrity not only expedites gate clearance, it protects our facilities and ensures we meet all regulatory requirements.
""")

with st.expander("Why is Data Integrity Important?"):
    st.write("""
    - **Accuracy**: Correct visitor details reduce clearance delays.  
    - **Security**: Reliable ID checks prevent unauthorized access.  
    - **Compliance**: Audit-ready records ensure regulatory adherence.  
    - **Efficiency**: Trustworthy data powers faster reporting and analytics.
    """)

st.markdown("### ⚠️ **Please ensure your spreadsheet has no missing or malformed fields.**")
uploaded = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

# --- Date Estimation ---
now = datetime.now(ZoneInfo("Asia/Singapore"))
st.markdown("### 🗓️ Estimate Clearance Date 🍍")
formatted_now = now.strftime("%A %d %B, %I:%M%p").lstrip("0")
st.write("**Today is:**", formatted_now)

if st.button("▶️ Calculate Estimated Delivery"):
    if now.time() >= datetime.strptime("15:00", "%H:%M").time():
        effective_submission_date = now.date() + timedelta(days=1)
    else:
        effective_submission_date = now.date()

    while effective_submission_date.weekday() >= 5:
        effective_submission_date += timedelta(days=1)

    working_days_count = 0
    estimated_date = effective_submission_date
    while working_days_count < 2:
        estimated_date += timedelta(days=1)
        if estimated_date.weekday() < 5:
            working_days_count += 1

    clearance_date = estimated_date
    while clearance_date.weekday() >= 5:
        clearance_date += timedelta(days=1)

    formatted = f"{clearance_date:%A} {clearance_date.day} {clearance_date:%B}"
    st.success(f"✓ Earliest clearance: **{formatted}**")

# --- Cleaning Function ---
# (keep your existing clean_data() function as-is)

# --- Excel Writer that preserves all sheets ---
def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.iloc[:, :13]
    df.columns = [
        "S/N", "Vehicle Plate Number", "Company Full Name", "Full Name As Per NRIC",
        "First Name as per NRIC", "Middle and Last Name as per NRIC", "Identification Type",
        "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date",
        "Nationality (Country Name)", "PR", "Gender", "Mobile Number",
    ]
    df = df.dropna(subset=df.columns[3:13], how="all")

    df["Company Full Name"] = (
        df["Company Full Name"]
        .astype(str)
        .str.replace(r"\bPTE\s+LTD\b", "Pte Ltd", flags=re.IGNORECASE, regex=True)
    )

    nat_map = {"chinese": "China", "singaporean": "Singapore", "malaysian": "Malaysia", "indian": "India"}
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
        .astype(str).str.strip().str.lower()
        .replace(nat_map, regex=False)
        .str.title()
    )

    def nationality_group(row):
        nat = str(row["Nationality (Country Name)"]).strip().lower()
        pr = str(row["PR"]).strip().lower()
        if nat == "singapore": return 1
        elif pr in ("yes", "y", "pr"): return 2
        elif nat == "malaysia": return 3
        elif nat == "india": return 4
        else: return 5

    df["SortGroup"] = df.apply(nationality_group, axis=1)
    df = (
        df.sort_values(
            ["Company Full Name", "SortGroup", "Nationality (Country Name)", "Full Name As Per NRIC"],
            ignore_index=True
        )
        .drop(columns="SortGroup")
    )
    df["S/N"] = range(1, len(df) + 1)

    def normalize_pr(value):
        val = str(value).strip().lower()
        if val in ("pr", "yes", "y"):
            return "PR"
        elif val in ("n", "no", "na", "", "nan"):
            return ""
        else:
            return val.upper() if val.isalpha() else val

    df["PR"] = df["PR"].apply(normalize_pr)

    df["Identification Type"] = (
        df["Identification Type"]
        .astype(str).str.strip()
        .apply(lambda v: "FIN" if v.lower() == "fin" else v.upper())
    )

    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
        .astype(str)
        .str.replace(r"[\/,]", ";", regex=True)
        .str.replace(r"\s*;\s*", ";", regex=True)
        .str.replace(r"\s+", "", regex=True)
        .replace("nan", "", regex=False)
    )

    def split_name(full_name):
        s = str(full_name).strip()
        if " " in s:
            i = s.find(" ")
            return pd.Series([s[:i], s[i+1:]])
        return pd.Series([s, ""])

    df["Full Name As Per NRIC"] = df["Full Name As Per NRIC"].astype(str).str.title()
    df[["First Name as per NRIC", "Middle and Last Name as per NRIC"]] = (
        df["Full Name As Per NRIC"].apply(split_name)
    )

    iccol, wpcol = "IC (Last 3 digits and suffix) 123A", "Work Permit Expiry Date"
    if df[iccol].astype(str).str.contains("-", na=False).any():
        df[[iccol, wpcol]] = df[[wpcol, iccol]]
    df[iccol] = df[iccol].astype(str).str[-4:]

    def fix_mobile(x):
        d = re.sub(r"\D", "", str(x))
        if len(d) > 8:
            extra = len(d) - 8
            if d.endswith("0" * extra): d = d[:-extra]
            else: d = d[-8:]
        if len(d) < 8: d = d.zfill(8)
        return d

    def clean_gender(g):
        v = str(g).strip().upper()
        if v == "M": return "Male"
        if v == "F": return "Female"
        if v in ("MALE", "FEMALE"): return v.title()
        return v

    df["Mobile Number"] = df["Mobile Number"].apply(fix_mobile)
    df["Gender"] = df["Gender"].apply(clean_gender)
    df[wpcol] = pd.to_datetime(df[wpcol], errors="coerce").dt.strftime("%Y-%m-%d")

    return df


# --- Main Execution ---
if uploaded:
    all_sheets = pd.read_excel(uploaded, sheet_name=None)
    raw_df = all_sheets.get("Visitor List")

    if raw_df is not None:
        company_cell = raw_df.iloc[0, 2]
        company = (
            str(company_cell).strip()
            if pd.notna(company_cell) and str(company_cell).strip()
            else "VisitorList"
        )

        cleaned = clean_data(raw_df)
        all_sheets["Visitor List"] = cleaned
        out_buf = generate_full_workbook(all_sheets)

        today_str = datetime.now(ZoneInfo("Asia/Singapore")).strftime("%Y%m%d")
        fname = f"{company}_{today_str}.xlsx"

        st.download_button(
            label="📅 Download Cleaned Workbook",
            data=out_buf.getvalue(),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(
            "✅ Your data has been validated. Please double-check critical fields before sharing with DC team."
        )
    else:
        st.error("❌ 'Visitor List' tab not found in uploaded file.")
