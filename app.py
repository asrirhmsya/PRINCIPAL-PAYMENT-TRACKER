import io
import re
from datetime import date, datetime, timedelta
from urllib.parse import urlparse

import pandas as pd
import plotly.express as px
import requests
import streamlit as st


st.set_page_config(
    page_title="Payment Tracker Dashboard",
    page_icon="💸",
    layout="wide",
)

st.title("💸 Payment Tracker Dashboard")
st.caption("AP MD Import | Principal Payment Monitoring | SharePoint Excel-based dashboard")


# =========================
# Helper functions
# =========================

def excel_serial_to_datetime(series: pd.Series) -> pd.Series:
    """Convert Excel serial dates or normal date strings into pandas datetime."""
    numeric = pd.to_numeric(series, errors="coerce")
    converted_numeric = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    converted_text = pd.to_datetime(series, errors="coerce")
    return converted_numeric.fillna(converted_text)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Clean column names so dashboard can handle small naming differences."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {
        "External Document No.": "Invoice No",
        "External Document No": "Invoice No",
        "Document No.": "Invoice No",
        "Document No": "Invoice No",
        "Original Amount": "Amount",
        "Vendor Name": "Vendor",
        "Currency Code": "Currency",
        "Company Code": "Company",
        "Brand Code": "Brand",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    for col in ["Posting Date", "Document Date", "Due Date"]:
        if col in df.columns:
            df[col] = excel_serial_to_datetime(df[col])

    if "Amount" in df.columns:
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)

    if "Due Date" in df.columns:
        today = pd.Timestamp(date.today())
        df["Days to Due"] = (df["Due Date"] - today).dt.days

        def due_bucket(x):
            if pd.isna(x):
                return "No Due Date"
            if x < 0:
                return "Overdue"
            if x <= 7:
                return "Due in 7 Days"
            if x <= 30:
                return "Due in 30 Days"
            return "Future Due"

        df["Due Status"] = df["Days to Due"].apply(due_bucket)
        df["Due Month Dashboard"] = df["Due Date"].dt.strftime("%b-%Y")
    else:
        df["Days to Due"] = None
        df["Due Status"] = "No Due Date"
        df["Due Month Dashboard"] = None

    if "PIC" not in df.columns:
        df["PIC"] = "Not Available"

    return df


def sharepoint_to_download_url(url: str) -> str:
    """Try to convert a SharePoint sharing URL into downloadable format."""
    if not url:
        return url
    if "download=1" in url:
        return url
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}download=1"


@st.cache_data(ttl=300, show_spinner=False)
def load_from_sharepoint(url: str) -> pd.DataFrame:
    download_url = sharepoint_to_download_url(url)
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(download_url, headers=headers, timeout=30)
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "")
    if "text/html" in content_type.lower():
        raise ValueError(
            "Link SharePoint masih membutuhkan login atau belum direct download. "
            "Untuk MVP, gunakan upload Excel manual dulu."
        )

    return pd.read_excel(io.BytesIO(response.content))


@st.cache_data(ttl=300, show_spinner=False)
def load_from_uploaded_file(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file)


# =========================
# Sidebar data source
# =========================

st.sidebar.header("Data Source")

source_mode = st.sidebar.radio(
    "Pilih sumber data:",
    ["Upload Excel", "SharePoint Link"],
    index=0,
)

raw_df = None

if source_mode == "Upload Excel":
    uploaded_file = st.sidebar.file_uploader(
        "Upload file Excel tracker",
        type=["xlsx", "xls"],
        help="Untuk MVP tercepat, upload Excel dari SharePoint ke sini.",
    )
    if uploaded_file is not None:
        raw_df = load_from_uploaded_file(uploaded_file)

else:
    sharepoint_url = st.sidebar.text_input(
        "Paste SharePoint Excel link",
        placeholder="https://...sharepoint.com/:x:/s/..."
    )
    if st.sidebar.button("Refresh from SharePoint", use_container_width=True):
        st.cache_data.clear()
    if sharepoint_url:
        try:
            raw_df = load_from_sharepoint(sharepoint_url)
        except Exception as e:
            st.error(f"Gagal membaca SharePoint link: {e}")
            st.info("Solusi cepat: download Excel dari SharePoint lalu pakai mode Upload Excel.")


if raw_df is None:
    st.info("Upload Excel dulu di sidebar kiri untuk menampilkan dashboard.")
    st.stop()


df = normalize_columns(raw_df)

required_cols = ["Due Date", "Vendor", "Amount", "Currency"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.warning(f"Kolom penting belum ditemukan: {', '.join(missing)}. Dashboard tetap ditampilkan sebisanya.")


# =========================
# Sidebar filters
# =========================

st.sidebar.header("Filters")

def multiselect_filter(label, col):
    if col in df.columns:
        options = sorted([x for x in df[col].dropna().unique()])
        return st.sidebar.multiselect(label, options=options, default=options)
    return None

selected_company = multiselect_filter("Company", "Company")
selected_brand = multiselect_filter("Brand", "Brand")
selected_currency = multiselect_filter("Currency", "Currency")
selected_vendor = multiselect_filter("Vendor", "Vendor")
selected_status = multiselect_filter("Due Status", "Due Status")

filtered = df.copy()

for col, selected in [
    ("Company", selected_company),
    ("Brand", selected_brand),
    ("Currency", selected_currency),
    ("Vendor", selected_vendor),
    ("Due Status", selected_status),
]:
    if selected is not None and col in filtered.columns:
        filtered = filtered[filtered[col].isin(selected)]


# =========================
# KPI cards
# =========================

today = pd.Timestamp(date.today())
end_7_days = today + pd.Timedelta(days=7)
end_30_days = today + pd.Timedelta(days=30)

total_invoice = len(filtered)
total_amount = filtered["Amount"].sum() if "Amount" in filtered.columns else 0
overdue_count = (filtered["Due Status"] == "Overdue").sum() if "Due Status" in filtered.columns else 0
due_7_count = (filtered["Due Status"] == "Due in 7 Days").sum() if "Due Status" in filtered.columns else 0
due_30_count = filtered["Due Status"].isin(["Due in 7 Days", "Due in 30 Days"]).sum() if "Due Status" in filtered.columns else 0

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
kpi1.metric("Total Invoice", f"{total_invoice:,}")
kpi2.metric("Total Amount", f"{total_amount:,.2f}")
kpi3.metric("Overdue", f"{overdue_count:,}")
kpi4.metric("Due ≤ 7 Days", f"{due_7_count:,}")
kpi5.metric("Due ≤ 30 Days", f"{due_30_count:,}")


# =========================
# Charts
# =========================

left, right = st.columns(2)

with left:
    st.subheader("Outstanding by Vendor")
    if "Vendor" in filtered.columns and "Amount" in filtered.columns and not filtered.empty:
        vendor_chart = (
            filtered.groupby("Vendor", as_index=False)["Amount"]
            .sum()
            .sort_values("Amount", ascending=False)
            .head(10)
        )
        fig = px.bar(vendor_chart, x="Amount", y="Vendor", orientation="h")
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=420)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Data vendor/amount belum tersedia.")

with right:
    st.subheader("Invoice Due Status")
    if "Due Status" in filtered.columns and not filtered.empty:
        status_chart = filtered["Due Status"].value_counts().reset_index()
        status_chart.columns = ["Due Status", "Invoice Count"]
        fig = px.pie(status_chart, names="Due Status", values="Invoice Count", hole=0.45)
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Data due status belum tersedia.")


st.subheader("Monthly Due Amount")
if "Due Month Dashboard" in filtered.columns and "Amount" in filtered.columns and not filtered.empty:
    month_chart = (
        filtered.dropna(subset=["Due Month Dashboard"])
        .groupby("Due Month Dashboard", as_index=False)["Amount"]
        .sum()
    )
    fig = px.bar(month_chart, x="Due Month Dashboard", y="Amount")
    fig.update_layout(height=360)
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Data due month belum tersedia.")


# =========================
# Detail table
# =========================

st.subheader("Invoice Detail")

default_cols = [
    "Posting Date",
    "Document Date",
    "Due Date",
    "Due Month",
    "Week",
    "Invoice No",
    "Vendor",
    "Company",
    "Brand",
    "Currency",
    "Amount",
    "Days to Due",
    "Due Status",
]

show_cols = [c for c in default_cols if c in filtered.columns]
if not show_cols:
    show_cols = filtered.columns.tolist()

display_df = filtered[show_cols].copy()

for col in ["Posting Date", "Document Date", "Due Date"]:
    if col in display_df.columns:
        display_df[col] = display_df[col].dt.strftime("%d-%b-%Y")

st.dataframe(display_df, use_container_width=True, hide_index=True)

csv = display_df.to_csv(index=False).encode("utf-8")
st.download_button(
    label="Download filtered data as CSV",
    data=csv,
    file_name="payment_tracker_filtered.csv",
    mime="text/csv",
)

st.caption("Note: SharePoint direct refresh depends on company permission/authentication setting. If blocked, use upload mode for MVP or set up Microsoft Graph API for secure auto-refresh.")