"""
Customer Feed Purchase Report
------------------------------
Reads sales data from a Google Sheet, maps customers to Zones using an
uploaded Customer List (Excel: Customer ID, Zone, ...), and shows a
report with:
    Customer Code | Customer Name | Last Feed Purchase Date |
    Due Date (Last Purchase + N days) | Last Order Date

Zone-wise slicer in the sidebar filters the whole report.

Run with:
    streamlit run app.py
"""

import io
import pandas as pd
import streamlit as st
from datetime import timedelta

st.set_page_config(page_title="Customer Feed Purchase Report", layout="wide")

# ----------------------------------------------------------------------
# CONFIG — edit these if your sheet changes
# ----------------------------------------------------------------------
SHEET_ID = "1S3csAE-E_hN8vstuHR0KkeAN7yCVQTFe4AkEVlw4vQw"
DEFAULT_GID = "0"                 # tab (gid) of the sales data sheet
FEED_PREFIX = "FEED"              # Item No. prefix that identifies "feed" items

# ----------------------------------------------------------------------
# DATA LOADERS
# ----------------------------------------------------------------------
@st.cache_data(ttl=300, show_spinner="Loading sales data from Google Sheet...")
def load_sales_data(sheet_id: str, gid: str) -> pd.DataFrame:
    """
    Loads the sales log from a published/shared Google Sheet.
    Requires sheet sharing: "Anyone with the link -> Viewer".
    """
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    df = pd.read_csv(url)

    df.columns = [c.strip() for c in df.columns]
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Customer Code"] = df["Customer Code"].astype(str).str.strip()
    df["Item No."] = df["Item No."].astype(str).str.strip()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Sales Amt"] = pd.to_numeric(df["Sales Amt"], errors="coerce").fillna(0)
    return df


@st.cache_data(ttl=300, show_spinner="Reading customer / zone list...")
def load_customer_master(file) -> pd.DataFrame:
    """
    Reads the uploaded Customer List Excel file.
    Expected columns (case-insensitive, flexible naming):
        Customer ID / Customer Code   -> customer code
        Zone                          -> zone
        Customer Name (optional)      -> fallback name
    """
    raw = pd.read_excel(file)
    raw.columns = [c.strip() for c in raw.columns]

    # Flexible column matching
    col_map = {}
    for c in raw.columns:
        cl = c.lower()
        if cl in ("customer id", "customer code", "cust id", "cust code"):
            col_map[c] = "Customer Code"
        elif cl == "zone":
            col_map[c] = "Zone"
        elif cl in ("customer name", "cust name", "name"):
            col_map[c] = "Customer Name (Master)"

    raw = raw.rename(columns=col_map)

    if "Customer Code" not in raw.columns or "Zone" not in raw.columns:
        raise ValueError(
            "Could not find 'Customer ID/Code' and 'Zone' columns in the "
            "uploaded file. Found columns: " + ", ".join(raw.columns)
        )

    raw["Customer Code"] = raw["Customer Code"].astype(str).str.strip()
    raw["Zone"] = raw["Zone"].astype(str).str.strip()

    keep_cols = ["Customer Code", "Zone"]
    if "Customer Name (Master)" in raw.columns:
        keep_cols.append("Customer Name (Master)")

    return raw[keep_cols].drop_duplicates(subset="Customer Code")


# ----------------------------------------------------------------------
# REPORT BUILDER
# ----------------------------------------------------------------------
def build_report(sales: pd.DataFrame, customers: pd.DataFrame, due_days: int) -> pd.DataFrame:
    sales = sales.merge(customers, on="Customer Code", how="left")

    # Prefer the name from the sales log; fall back to master list name
    if "Customer Name (Master)" in sales.columns:
        sales["Customer Name"] = sales["Customer Name"].fillna(sales["Customer Name (Master)"])

    is_feed = sales["Item No."].str.upper().str.startswith(FEED_PREFIX)

    # Last order overall (any item)
    last_order = (
        sales.groupby("Customer Code")["Date"].max().rename("Last Order")
    )

    # Last feed purchase only
    feed_sales = sales[is_feed]
    last_feed = (
        feed_sales.groupby("Customer Code")["Date"].max().rename("Last Feed Purchase Date")
    )

    # Static per-customer info (name, zone) — take the latest non-null row
    info = (
        sales.sort_values("Date")
        .groupby("Customer Code")
        .agg({"Customer Name": "last", "Zone": "last"})
    )

    report = info.join(last_feed).join(last_order).reset_index()

    report["Due date last Purchase"] = report["Last Feed Purchase Date"] + timedelta(days=due_days)

    today = pd.Timestamp.now().normalize()

    def remark(d):
        if pd.isna(d):
            return "No Feed Purchase"
        elif d < today:
            return "Overdue"
        elif d <= today + timedelta(days=7):
            return "Due Soon"
        return "OK"

    report["Remarks"] = report["Due date last Purchase"].apply(remark)

    for col in ["Last Feed Purchase Date", "Due date last Purchase", "Last Order"]:
        report[col] = report[col].dt.strftime("%Y-%m-%d")

    # Zone kept in the dataframe (used for filtering) but not shown in the final table
    report = report[
        ["Customer Code", "Customer Name", "Zone", "Last Feed Purchase Date",
         "Due date last Purchase", "Remarks", "Last Order"]
    ]
    return report.sort_values("Customer Code").reset_index(drop=True)


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
st.title("📦 Customer Feed Purchase Report")

with st.sidebar:
    st.header("⚙️ Settings")
    gid = st.text_input("Sales sheet tab (gid)", value=DEFAULT_GID)
    due_days = st.number_input("Due after last feed purchase (days)", min_value=1, value=30, step=1)

    st.markdown("---")
    st.subheader("Customer / Zone List")
    customer_file = st.file_uploader(
        "Upload Customer List (.xlsx) with Customer ID + Zone", type=["xlsx", "xls"]
    )

if customer_file is None:
    st.info("⬅️ Upload your Customer List Excel file (Customer ID + Zone columns) in the sidebar to continue.")
    st.stop()

try:
    sales_df = load_sales_data(SHEET_ID, gid)
    customers_df = load_customer_master(customer_file)
except Exception as e:
    st.error(f"Error loading data: {e}")
    st.stop()

report_df = build_report(sales_df, customers_df, due_days)

# Zone-wise slicer — table only appears after a zone is picked
zones = sorted([z for z in report_df["Zone"].dropna().unique() if z and z != "nan"])

st.subheader("🌍 Select Zone")
selected_zones = st.multiselect("Zone", options=zones, placeholder="Choose one or more zones...")

if not selected_zones:
    st.info("👆 Select at least one zone above to display the report.")
    st.stop()

filtered = report_df[report_df["Zone"].isin(selected_zones)]

DISPLAY_COLS = ["Customer Code", "Customer Name", "Last Feed Purchase Date",
                "Due date last Purchase", "Remarks", "Last Order"]
table = filtered[DISPLAY_COLS]

# ----------------------------------------------------------------------
# DISPLAY
# ----------------------------------------------------------------------
col1, col2, col3 = st.columns(3)
col1.metric("Customers Shown", len(table))
col2.metric("Overdue", (table["Remarks"] == "Overdue").sum())
col3.metric("Due Soon", (table["Remarks"] == "Due Soon").sum())

st.markdown("---")


def highlight_remarks(row):
    if row["Remarks"] == "Overdue":
        return ["background-color: #ffcccc"] * len(row)
    elif row["Remarks"] == "Due Soon":
        return ["background-color: #fff3cd"] * len(row)
    return [""] * len(row)


st.dataframe(
    table.style.apply(highlight_remarks, axis=1),
    use_container_width=True,
    hide_index=True,
)

csv = table.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download Report as CSV", data=csv, file_name="feed_purchase_report.csv", mime="text/csv")
