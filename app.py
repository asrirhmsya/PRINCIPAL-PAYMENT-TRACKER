import io
from datetime import date

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
st.caption("AP MD Import | Principal Payment Monitoring | Weekly Payment Planning")


# =========================
# Helper functions
# =========================

def excel_serial_to_datetime(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    converted_numeric = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    converted_text = pd.to_datetime(series, errors="coerce")
    return converted_numeric.fillna(converted_text)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
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

        # Weekly planning columns
        # Prefer Week column from Excel because AP/Treasury already maintains payment week manually.
        if "Week" in df.columns:
            df["Payment Week"] = df["Week"].astype(str).str.strip()
        else:
            iso = df["Due Date"].dt.isocalendar()
            df["Due Year"] = iso["year"]
            df["Due Week No"] = iso["week"]
            df["Payment Week"] = "W" + df["Due Week No"].astype(str).str.zfill(2) + " - " + df["Due Year"].astype(str)

        week_start = df["Due Date"] - pd.to_timedelta(df["Due Date"].dt.weekday, unit="D")
        week_end = week_start + pd.Timedelta(days=6)
        df["Week Range"] = week_start.dt.strftime("%d %b") + " - " + week_end.dt.strftime("%d %b %Y")
    else:
        df["Days to Due"] = None
        df["Due Status"] = "No Due Date"
        df["Due Month Dashboard"] = None
        df["Payment Week"] = "No Due Date"
        df["Week Range"] = "No Due Date"

    if "Week" in df.columns:
        df["Source Week"] = df["Week"].astype(str)
    else:
        df["Source Week"] = df["Payment Week"]

    if "PIC" not in df.columns:
        df["PIC"] = "Not Available"

    return df


def sharepoint_to_download_url(url: str) -> str:
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
# Data source
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


# =========================
# Filters
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
selected_week = multiselect_filter("Payment Week", "Payment Week")

filtered = df.copy()

for col, selected in [
    ("Company", selected_company),
    ("Brand", selected_brand),
    ("Currency", selected_currency),
    ("Vendor", selected_vendor),
    ("Due Status", selected_status),
    ("Payment Week", selected_week),
]:
    if selected is not None and col in filtered.columns:
        filtered = filtered[filtered[col].isin(selected)]


# =========================
# KPI cards
# =========================

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
# Weekly Payment Summary
# =========================

st.header("📅 Weekly Payment Summary")

if "Payment Week" in filtered.columns and "Amount" in filtered.columns and not filtered.empty:
    weekly_summary = (
        filtered
        .groupby(["Payment Week", "Week Range", "Currency"], dropna=False)
        .agg(
            Invoice_Count=("Amount", "count"),
            Total_Amount=("Amount", "sum"),
            Vendor_Count=("Vendor", "nunique") if "Vendor" in filtered.columns else ("Amount", "count")
        )
        .reset_index()
    )

    week_total = (
        filtered
        .groupby(["Payment Week", "Week Range"], dropna=False)
        .agg(
            Invoice_Count=("Amount", "count"),
            Total_Amount=("Amount", "sum"),
        )
        .reset_index()
    )

    def week_sort_value(x):
        import re
        m = re.search(r"\d+", str(x))
        return int(m.group()) if m else 9999

    weekly_summary["_week_sort"] = weekly_summary["Payment Week"].apply(week_sort_value)
    week_total["_week_sort"] = week_total["Payment Week"].apply(week_sort_value)
    weekly_summary = weekly_summary.sort_values(["_week_sort", "Currency"]).drop(columns=["_week_sort"])
    week_total = week_total.sort_values("_week_sort").drop(columns=["_week_sort"])

    c1, c2 = st.columns([1.2, 1])

    with c1:
        st.subheader("Payment Plan by Week")
        fig = px.bar(
            week_total,
            x="Payment Week",
            y="Total_Amount",
            text="Invoice_Count",
            hover_data=["Week Range", "Invoice_Count"],
        )
        fig.update_layout(height=380, xaxis_title="Payment Week", yaxis_title="Total Amount")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Weekly Amount by Currency")
        fig = px.bar(
            weekly_summary,
            x="Payment Week",
            y="Total_Amount",
            color="Currency",
            hover_data=["Week Range", "Invoice_Count", "Vendor_Count"],
        )
        fig.update_layout(height=380, xaxis_title="Payment Week", yaxis_title="Total Amount")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Weekly Summary Table")
    st.dataframe(
        weekly_summary.rename(columns={
            "Invoice_Count": "Invoice Count",
            "Total_Amount": "Total Amount",
            "Vendor_Count": "Vendor Count",
        }),
        use_container_width=True,
        hide_index=True
    )

    available_weeks = week_total["Payment Week"].tolist()
    selected_detail_week = st.selectbox("Lihat detail invoice untuk week:", available_weeks)

    detail_week_df = filtered[filtered["Payment Week"] == selected_detail_week].copy()

    st.subheader(f"Invoice Detail - {selected_detail_week}")
    detail_cols = [
        "Due Date", "Week Range", "Invoice No", "Vendor", "Company", "Brand",
        "Currency", "Amount", "Days to Due", "Due Status"
    ]
    detail_cols = [c for c in detail_cols if c in detail_week_df.columns]

    display_week = detail_week_df[detail_cols].copy()
    if "Due Date" in display_week.columns:
        display_week["Due Date"] = display_week["Due Date"].dt.strftime("%d-%b-%Y")

    st.dataframe(display_week, use_container_width=True, hide_index=True)

else:
    st.info("Data weekly summary belum tersedia.")


# =========================
# Other charts
# =========================

st.header("📊 Dashboard Overview")

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

st.header("📋 Full Invoice Detail")

default_cols = [
    "Posting Date",
    "Document Date",
    "Due Date",
    "Payment Week",
    "Week Range",
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

st.caption("Refresh MVP: update Excel di SharePoint, download file terbaru, lalu upload ulang ke dashboard. Direct SharePoint refresh bisa jadi phase berikutnya.")
