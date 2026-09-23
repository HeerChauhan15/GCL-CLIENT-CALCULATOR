import streamlit as st
import pandas as pd
import numpy as np
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Aviva Client Premium Calculator",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Aviva Client Premium Calculator")
st.markdown("Select plan details below")

# Single Life only
FILE_MAP = {
    ("Level",    "Home Loan"):     "Single-Homeloan-Level.xlsx",
    ("Level",    "LAP"):           "Single-Lap-Level.xlsx",
    ("Reducing", "Home Loan"):     "Single-Homeloan-Reducing.xlsx",
    ("Reducing", "LAP"):           "Single-Lap-Reducing.xlsx",
    ("Level",    "Personal Loan"): "Single-Personal-Loan- Level.xlsx",
    ("Reducing", "Personal Loan"): "Single-Personal-Loan- Reducing.xlsx",
}

# GST is fixed and always applied on top of the Loader-adjusted rate.
GST_RATE_FIXED = 18.0

# ============================================
# LOADER + GST FORMULA (applied to every rate before premium is computed):
#   Rate After Loader   = Base Rate / (1 - Loader% / 100)
#   Premium (Excl. GST) = Rate After Loader x (Sum Assured / 1,00,000)
#   GST Amount          = Premium (Excl. GST) x (GST% / 100)
#   Premium (Incl. GST) = Premium (Excl. GST) + GST Amount
# ============================================
def compute_premium_breakdown(base_rate, loader_pct, sum_assured, gst_pct=GST_RATE_FIXED):
    after_loader = base_rate / (1 - (loader_pct / 100.0))
    premium_excl_gst = after_loader * (sum_assured / 100000.0)
    gst_amount = premium_excl_gst * (gst_pct / 100.0)
    premium_incl_gst = premium_excl_gst + gst_amount
    return premium_excl_gst, gst_amount, premium_incl_gst


def load_rate_table(cover_type, loan_type):
    fname = FILE_MAP[(cover_type, loan_type)]
    if not os.path.exists(fname):
        raise FileNotFoundError(
            f"File not found: '{fname}' — Please make sure this file is in the GitHub repo."
        )

    raw = pd.read_excel(fname, sheet_name="Sheet1", header=None)

    header_row = None
    for i, row in raw.iterrows():
        for val in row.values:
            if isinstance(val, str) and "AGE" in val.upper():
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        raise ValueError("Could not find AGE/TERM header row in the file.")

    df = pd.read_excel(fname, sheet_name="Sheet1", header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    age_col = df.columns[0]
    df = df.dropna(subset=[age_col])
    df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
    df = df.dropna(subset=[age_col])
    df[age_col] = df[age_col].astype(int)
    df = df.set_index(age_col)

    tenure_map = {}
    for col in df.columns:
        try:
            tenure_map[int(float(col))] = col
        except Exception:
            pass

    return df, tenure_map


def get_rate(df, tenure_map, age, tenure):
    if age not in df.index:
        raise ValueError(f"Age {age} not found in rate table.")
    if tenure not in tenure_map:
        raise ValueError(f"Tenure {tenure} yrs not found in rate table.")
    return float(df.loc[age, tenure_map[tenure]])


def find_column(df, target):
    """Exact (case/space-insensitive) match."""
    target_norm = target.strip().lower().replace(" ", "")
    for col in df.columns:
        col_norm = str(col).strip().lower().replace(" ", "")
        if col_norm == target_norm:
            return col
    return None


def find_sum_assured_columns(df):
    """
    Detects the Sum Assured/Sum Insured column AND the Loan Outstanding column
    (if both are present). Per-row logic then uses Sum Assured when that row
    has a value, and falls back to Loan Outstanding for that row only when
    Sum Assured is blank/missing.
    """
    sa_col = None
    lo_col = None
    for col in df.columns:
        norm = re.sub(r'[\s_\-/]+', '', str(col).lower())
        if sa_col is None and ('sumassured' in norm or 'suminsured' in norm):
            sa_col = col
        if lo_col is None and ('loanoutstanding' in norm or 'outstandingamount' in norm or 'outstandingloan' in norm
                or norm == 'outstanding' or 'loanos' in norm or norm == 'os'):
            lo_col = col
    return sa_col, lo_col


# ============================================
# DROPDOWNS
# ============================================

col1, col2 = st.columns(2)
with col1:
    loan_type = st.selectbox("Select Loan Type", ["Home Loan", "LAP", "Personal Loan"])
with col2:
    cover_type = st.selectbox("Select Type of Cover", ["Level", "Reducing"])

# ============================================
# SHARED LOADER % — applied to every rate (Manual + Bulk)
# before computing premium. GST @ 18% is then added automatically on top.
# ============================================
st.subheader("⚙️ Loader Setting")
loader_pct_input = st.number_input(
    "Loader % (optional; GST 18% added automatically)",
    min_value=0.0,
    max_value=99.99,
    value=None,
    step=1.0,
    placeholder="Enter loader % (optional, defaults to 0)",
    key="shared_loader_pct"
)
loader_pct = loader_pct_input if loader_pct_input is not None else 0.0

# ============================================
# SUM ASSURED RANGE — rates in the backend files are per ₹1,00,000
# Home Loan: ₹50,000 – ₹60,00,000 | LAP: ₹50,000 – ₹40,00,000
# Personal Loan: ₹25,000 – ₹30,00,000
# ============================================
if loan_type == "Home Loan":
    sa_min, sa_max = 50000, 6000000
elif loan_type == "LAP":
    sa_min, sa_max = 50000, 4000000
else:  # Personal Loan
    sa_min, sa_max = 25000, 3000000

st.divider()

# ============================================
# MANUAL SECTION
# ============================================

st.subheader("🔢 Manual Rate Lookup")

if loan_type == "Home Loan":
    min_tenure, max_tenure = 2, 25
elif loan_type == "LAP":
    min_tenure, max_tenure = 2, 10
else:  # Personal Loan
    min_tenure, max_tenure = 2, 5

col3, col4 = st.columns(2)
with col3:
    age = st.number_input("Enter Age", min_value=18, max_value=65, value=30, step=1)
with col4:
    tenure = st.number_input(
        "Enter Tenure",
        min_value=min_tenure,
        max_value=max_tenure,
        value=min_tenure,
        step=1
    )
    st.caption("📅 Tenure is in Years")

sum_assured_manual = st.number_input(
    "Select Sum Assured (₹)",
    min_value=sa_min,
    max_value=sa_max,
    value=sa_min,
    step=1,
    help=f"Enter the exact Sum Assured for this member (₹{sa_min:,} – ₹{sa_max:,}).",
    key="sa_manual"
)

if st.button("Get Rate", type="primary"):
    try:
        df_rates, tenure_map = load_rate_table(cover_type, loan_type)
        base_rate = get_rate(df_rates, tenure_map, age, tenure)
        premium_excl_gst, gst_amount, premium_incl_gst = compute_premium_breakdown(
            base_rate, loader_pct, sum_assured_manual
        )
        st.success(
            f"✅ {loan_type} | {cover_type} Cover | Age {age} | Tenure {tenure} yrs | "
            f"Sum Assured ₹{sum_assured_manual:,} | Loader {loader_pct}% | GST {GST_RATE_FIXED}%"
        )
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("Premium (Excl. GST)", f"₹ {round(premium_excl_gst):,}")
        with m2:
            st.metric(f"GST Amount ({GST_RATE_FIXED:.0f}%)", f"₹ {gst_amount:,.2f}")
        with m3:
            st.metric("Premium (Incl. GST)", f"₹ {premium_incl_gst:,.2f}")
    except Exception as e:
        st.error(f"Error: {e}")

st.divider()

# ============================================
# EXCEL UPLOAD SECTION
# ============================================

st.subheader("📂 Upload Member Data for Bulk Rate Lookup")

st.markdown(
    "Excel needs: **Name**, **Age**, **Sum Assured** (or **Loan Outstanding**), **Tenure** (years)."
)

st.caption(f"Uses each row's own Sum Assured (₹{sa_min:,} – ₹{sa_max:,}).")
st.warning("⚠️ Please make sure you have selected **Loan Type** and **Type of Cover** above before uploading your Excel file.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        if loan_type == "Home Loan":
            min_t, max_t = 2, 25
        elif loan_type == "LAP":
            min_t, max_t = 2, 10
        else:  # Personal Loan
            min_t, max_t = 2, 5

        df_rates, tenure_map = load_rate_table(cover_type, loan_type)

        name_col = find_column(df, "Name")
        age_col = find_column(df, "Age")
        tenure_col = find_column(df, "Tenure")
        sa_col, lo_col = find_sum_assured_columns(df)

        if not name_col or not age_col or not tenure_col:
            raise ValueError(
                "Excel must contain mandatory columns: Name, Age, and Tenure."
            )
        if not sa_col and not lo_col:
            raise ValueError("Excel must contain a Sum Assured column (e.g. 'Sum Assured', 'Sum Insured') or, failing that, a 'Loan Outstanding' column.")

        df[age_col] = pd.to_numeric(df[age_col], errors='coerce')
        df[tenure_col] = pd.to_numeric(df[tenure_col], errors='coerce')
        if sa_col:
            df[sa_col] = pd.to_numeric(df[sa_col], errors='coerce')
        if lo_col:
            df[lo_col] = pd.to_numeric(df[lo_col], errors='coerce')

        if df[tenure_col].dropna().median() > 30:
            st.info("ℹ️ Tenure values look like months — auto-converting to years.")
            df[tenure_col] = (df[tenure_col] / 12).round(0).astype('Int64')
        else:
            df[tenure_col] = df[tenure_col].round(0).astype('Int64')

        df[age_col] = df[age_col].round(0).astype('Int64')

        premiums_excl_gst = []
        gst_amounts = []
        premiums_incl_gst = []
        statuses = []
        sa_used_list = []
        for idx, row in df.iterrows():
            try:
                r_age = int(row[age_col]) if pd.notna(row[age_col]) else None
                r_tenure = int(row[tenure_col]) if pd.notna(row[tenure_col]) else None

                # Per-row: use Sum Assured if this row has it; otherwise fall back
                # to Loan Outstanding for this row.
                r_sa = None
                if sa_col and pd.notna(row[sa_col]):
                    r_sa = float(row[sa_col])
                elif lo_col and pd.notna(row[lo_col]):
                    r_sa = float(row[lo_col])

                if r_sa is None:
                    raise ValueError("Sum Assured / Loan Outstanding value missing")

                if r_age is None or r_age < 18 or r_age > 65:
                    raise ValueError("Age must be between 18 and 65")
                if r_tenure is None or r_tenure < min_t or r_tenure > max_t:
                    raise ValueError(f"Tenure must be between {min_t} and {max_t} yrs")
                if r_sa < sa_min or r_sa > sa_max:
                    raise ValueError(f"Sum Assured must be between ₹{sa_min:,} and ₹{sa_max:,}")

                r_base = get_rate(df_rates, tenure_map, r_age, r_tenure)
                r_excl, r_gst, r_incl = compute_premium_breakdown(r_base, loader_pct, r_sa)
                premiums_excl_gst.append(round(r_excl))
                gst_amounts.append(round(r_gst, 2))
                premiums_incl_gst.append(round(r_incl, 2))
                statuses.append("✅")
                sa_used_list.append(r_sa)
            except Exception as e:
                premiums_excl_gst.append(None)
                gst_amounts.append(None)
                premiums_incl_gst.append(None)
                statuses.append(f"❌ {e}")
                sa_used_list.append(None)

        df["Sum Assured Used"] = sa_used_list
        df["Premium (Excl. GST)"] = premiums_excl_gst
        df[f"GST Amount ({GST_RATE_FIXED:.0f}%)"] = gst_amounts
        df["Premium (Incl. GST)"] = premiums_incl_gst
        df["Status"] = statuses

        display_sa_col = sa_col if sa_col else lo_col
        core_cols = [
            name_col, age_col, tenure_col, display_sa_col, "Sum Assured Used",
            "Premium (Excl. GST)", f"GST Amount ({GST_RATE_FIXED:.0f}%)", "Premium (Incl. GST)"
        ]
        extra_cols = [c for c in df.columns if c not in core_cols]
        df = df[core_cols + extra_cols]

        total_excl_gst = pd.to_numeric(pd.Series(premiums_excl_gst), errors='coerce').sum()
        total_gst = pd.to_numeric(pd.Series(gst_amounts), errors='coerce').sum()
        total_incl_gst = pd.to_numeric(pd.Series(premiums_incl_gst), errors='coerce').sum()

        t1, t2, t3 = st.columns(3)
        with t1:
            st.metric("💰 Total Premium (Excl. GST)", f"₹ {round(total_excl_gst):,}")
        with t2:
            st.metric(f"💰 Total GST ({GST_RATE_FIXED:.0f}%)", f"₹ {total_gst:,.2f}")
        with t3:
            st.metric("💰 Grand Total Premium (Incl. GST)", f"₹ {total_incl_gst:,.2f}")

        st.subheader("Rate Lookup Output")
        st.dataframe(df, use_container_width=True)

        total_row = {c: "" for c in df.columns}
        total_row[name_col] = "TOTAL PREMIUM"
        total_row["Premium (Excl. GST)"] = round(total_excl_gst)
        total_row[f"GST Amount ({GST_RATE_FIXED:.0f}%)"] = round(total_gst, 2)
        total_row["Premium (Incl. GST)"] = round(total_incl_gst, 2)
        df_out = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

        output_file = "Rate_Output.xlsx"
        df_out.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="⬇ Download Output Excel",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
