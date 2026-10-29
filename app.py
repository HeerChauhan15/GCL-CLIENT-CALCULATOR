import streamlit as st
import pandas as pd
import numpy as np
import os
import re

st.set_page_config(
    page_title="GCL Client Calculator",
    page_icon="💰",
    layout="wide"
)

st.title("💰 GCL Client Calculator")
st.markdown("Select plan details below")

# GST is fixed and always applied on top of the Loader-adjusted rate.
GST_RATE_FIXED = 18.0

# ============================================
# SEGMENT CONFIG
# One place to control files, Sum Assured range and Tenure range per loan type.
# "Level" cover for Vehicle Loan uses VL-Flat.xlsx.
# ============================================
SEGMENTS = {
    "Home Loan": {
        "files": {"Level": "Single-Homeloan-Level.xlsx", "Reducing": "Single-Homeloan-Reducing.xlsx"},
        "sa_min": 50000, "sa_max": 6000000, "t_min": 5, "t_max": 25,
    },
    "Lap": {
        "files": {"Level": "Single-Lap-Level.xlsx", "Reducing": "Single-Lap-Reducing.xlsx"},
        "sa_min": 50000, "sa_max": 4000000, "t_min": 2, "t_max": 10,
    },
    "Personal Loan": {
        "files": {"Level": "Single-Personal-Loan- Level.xlsx", "Reducing": "Single-Personal-Loan- Reducing.xlsx"},
        "sa_min": 25000, "sa_max": 3000000, "t_min": 2, "t_max": 5,
    },
    "Business Loan": {
        "files": {"Level": "BL- Level.xlsx", "Reducing": "BL- Reducing.xlsx"},
        "sa_min": 50000, "sa_max": 5000000, "t_min": 2, "t_max": 10,
    },
    "Car Loan": {
        "files": {"Level": "CL- Level.xlsx", "Reducing": "CL-Reducing.xlsx"},
        "sa_min": 100000, "sa_max": 5000000, "t_min": 2, "t_max": 7,
    },
    "Education Loan": {
        "files": {"Level": "EL- Level.xlsx", "Reducing": "EL- Reducing.xlsx"},
        "sa_min": 25000, "sa_max": 5000000, "t_min": 2, "t_max": 10,
    },
    "Vehicle Loan": {
        "files": {"Level": "VL-Flat.xlsx", "Reducing": "VL-Reducing.xlsx"},
        "sa_min": 25000, "sa_max": 500000, "t_min": 2, "t_max": 5,
    },
}

AGE_MIN, AGE_MAX = 18, 65


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


def _norm_name(s):
    return re.sub(r"\s+", "", str(s)).lower()


def resolve_file(fname):
    """Exact match first; otherwise match ignoring spaces/case, so small
    filename differences (e.g. 'CL-Reducing' vs 'CL- Reducing') don't break."""
    if os.path.exists(fname):
        return fname
    target = _norm_name(fname)
    for f in os.listdir("."):
        if _norm_name(f) == target:
            return f
    return None


def load_rate_table(cover_type, loan_type):
    fname = SEGMENTS[loan_type]["files"][cover_type]
    path = resolve_file(fname)
    if path is None:
        raise FileNotFoundError(
            f"File not found: '{fname}' — Please make sure this file is in the GitHub repo."
        )

    raw = pd.read_excel(path, sheet_name="Sheet1", header=None)

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

    df = pd.read_excel(path, sheet_name="Sheet1", header=header_row)
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
    loan_type = st.selectbox("Select Loan Type", list(SEGMENTS.keys()))
with col2:
    cover_type = st.selectbox("Select Type of Cover", ["Level", "Reducing"])

cfg = SEGMENTS[loan_type]
sa_min, sa_max = cfg["sa_min"], cfg["sa_max"]
min_tenure, max_tenure = cfg["t_min"], cfg["t_max"]

# Compact summary of the active selection and its limits
st.caption(
    f"**{loan_type} · {cover_type} cover** — "
    f"Sum Assured ₹{sa_min:,} – ₹{sa_max:,} · Tenure {min_tenure} – {max_tenure} yrs · "
    f"Age {AGE_MIN} – {AGE_MAX}"
)

# ============================================
# LOADER — applied to every rate (Manual + Bulk). GST @ 18% added on top.
# Explanation lives in the (?) help tooltip instead of cluttering the page.
# ============================================
st.subheader("⚙️ Loader")
loader_col, _ = st.columns([1, 2])
with loader_col:
    loader_pct_input = st.number_input(
        "Loader %",
        min_value=0.0,
        max_value=99.99,
        value=None,
        step=1.0,
        placeholder="Optional (defaults to 0)",
        help="18% GST is added automatically on top of the premium.",
        key="shared_loader_pct"
    )
loader_pct = loader_pct_input if loader_pct_input is not None else 0.0

st.divider()

# ============================================
# MANUAL SECTION
# (keys include loan_type so ranges reset cleanly when the segment changes)
# ============================================
st.subheader("🔢 Individual Premium Calculator")

col3, col4 = st.columns(2)
with col3:
    age = st.number_input(
        "Enter Age", min_value=AGE_MIN, max_value=AGE_MAX, value=30, step=1,
        key=f"age_manual_{loan_type}"
    )
with col4:
    tenure = st.number_input(
        "Enter Tenure (Years)",
        min_value=min_tenure,
        max_value=max_tenure,
        value=min_tenure,
        step=1,
        help=f"Tenure in years ({min_tenure} – {max_tenure}).",
        key=f"tenure_manual_{loan_type}"
    )

sum_assured_manual = st.number_input(
    "Select Sum Assured (₹)",
    min_value=sa_min,
    max_value=sa_max,
    value=sa_min,
    step=1,
    help=f"Enter the exact Sum Assured for this member (₹{sa_min:,} – ₹{sa_max:,}).",
    key=f"sa_manual_{loan_type}"
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
head_col, info_col = st.columns([4, 1])
with head_col:
    st.subheader("📂 Group Premium Calculator")
with info_col:
    # Popover = the "popup" for Excel requirements; keeps the page clean.
    with st.popover("ℹ️ Excel format", use_container_width=True):
        st.markdown("**Required columns**")
        st.markdown(
            "- **Name**\n"
            "- **Age**\n"
            "- **Sum Assured** (or **Loan Outstanding** if blank)\n"
            "- **Tenure** (in years)"
        )
        st.markdown("**Limits for current selection**")
        st.markdown(
            f"- Loan Type: **{loan_type}** · Cover: **{cover_type}**\n"
            f"- Sum Assured: ₹{sa_min:,} – ₹{sa_max:,}\n"
            f"- Tenure: {min_tenure} – {max_tenure} yrs\n"
            f"- Age: {AGE_MIN} – {AGE_MAX}"
        )
        st.caption("Each row uses its own Sum Assured. Rows outside the limits show an ❌ status.")

uploaded_file = st.file_uploader(
    f"Upload Excel File for {loan_type} · {cover_type} cover",
    type=["xlsx"],
    help="Rates are picked based on the Loan Type and Type of Cover selected above."
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]

        st.subheader("Uploaded Data Preview")
        st.dataframe(df.head())

        min_t, max_t = min_tenure, max_tenure

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

                if r_age is None or r_age < AGE_MIN or r_age > AGE_MAX:
                    raise ValueError(f"Age must be between {AGE_MIN} and {AGE_MAX}")
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
