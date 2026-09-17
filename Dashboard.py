import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(
    page_title="Healthcare Dashboard",
    page_icon="🏥",
    layout="wide",
)

st.title("🏥 Healthcare Data Dashboard")
st.caption("Multi-clinic overview · Auto-loaded from your automation output")

# ---------- Load the output file ----------
@st.cache_data(ttl=60)  # refresh every 60s
@st.cache_data(ttl=60)
def load_data():
    # Try known names/locations
    candidates = [
        Path("Healthcare_data_Automaion.xlsx"),   # ← your actual file
        Path("Healthcare_Data_Automation.xlsx"),
        Path("output/Healthcare_Data_Automation.xlsx"),
        Path("output/Healthcare_data_Automaion.xlsx"),
    ]
    # Fallback: search anywhere in the project
    candidates += list(Path(".").rglob("*.xlsx"))

    for p in candidates:
        if p.exists():
            try:
                master = pd.read_excel(p, sheet_name=0)
                try:
                    summary = pd.read_excel(p, sheet_name=1)
                except Exception:
                    summary = None
                st.sidebar.success(f"📂 Loaded: {p.name}")
                return master, summary
            except Exception as e:
                st.sidebar.warning(f"⚠️ Skipped {p}: {e}")

    return None

data = load_data()

if data is None:
    st.warning("⚠️ No output file found. Run `python main.py` first to generate `output/Healthcare_Data_Automation.xlsx`.")
    st.stop()

master, summary = data

# Ensure proper dtypes
if "Date" in master.columns:
    master["Date"] = pd.to_datetime(master["Date"], errors="coerce")

# ---------- Sidebar Filters ----------
st.sidebar.header("🔍 Filters")

if "Location" in master.columns:
    locations = master["Location"].dropna().unique().tolist()
    picked_loc = st.sidebar.multiselect("Clinic", locations, default=locations)
    if picked_loc:
        master = master[master["Location"].isin(picked_loc)]

if "Doctor" in master.columns:
    doctors = master["Doctor"].dropna().unique().tolist()
    picked_doc = st.sidebar.multiselect("Doctor", doctors, default=doctors)
    if picked_doc:
        master = master[master["Doctor"].isin(picked_doc)]

if "Service_Type" in master.columns:
    services = master["Service_Type"].dropna().unique().tolist()
    picked_svc = st.sidebar.multiselect("Service", services, default=services)
    if picked_svc:
        master = master[master["Service_Type"].isin(picked_svc)]

if "Status" in master.columns:
    statuses = master["Status"].dropna().unique().tolist()
    picked_status = st.sidebar.multiselect("Status", statuses, default=statuses)
    if picked_status:
        master = master[master["Status"].isin(picked_status)]

if "Date" in master.columns and master["Date"].notna().any():
    min_d = master["Date"].min().date()
    max_d = master["Date"].max().date()
    date_range = st.sidebar.date_input("Date range", [min_d, max_d])
    if len(date_range) == 2:
        master = master[
            (master["Date"].dt.date >= date_range[0]) &
            (master["Date"].dt.date <= date_range[1])
        ]

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing **{len(master):,}** records")

# ---------- KPI Row ----------
st.subheader("📌 Key Metrics")
c1, c2, c3, c4 = st.columns(4)

total_amount = master["Amount"].sum() if "Amount" in master else 0
total_patients = master["Patient_Name"].nunique() if "Patient_Name" in master else len(master)
avg_sale = master["Amount"].mean() if "Amount" in master else 0
paid_pct = (
    (master["Status"] == "Paid").mean() * 100
    if "Status" in master else 0
)

c1.metric("💰 Total Revenue", f"${total_amount:,.0f}")
c2.metric("👥 Total Patients", f"{total_patients:,}")
c3.metric("📈 Avg Sale", f"${avg_sale:,.2f}")
c4.metric("✅ Paid %", f"{paid_pct:.1f}%")

# ---------- Tabs ----------
tab1, tab2, tab3, tab4 = st.tabs(
    ["📊 Overview", "🏥 By Clinic", "👨‍⚕️ By Doctor", "📋 Raw Data"]
)

# ----- Tab 1: Overview -----
with tab1:
    col1, col2 = st.columns(2)

    with col1:
        if "Date" in master.columns and "Amount" in master.columns:
            trend = (
                master.groupby(master["Date"].dt.date)["Amount"]
                .sum().reset_index()
            )
            trend.columns = ["Date", "Amount"]
            fig = px.line(
                trend, x="Date", y="Amount",
                title="💰 Revenue Over Time",
                markers=True,
            )
            fig.update_layout(hovermode="x unified")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Service_Type" in master.columns and "Amount" in master.columns:
            svc = master.groupby("Service_Type")["Amount"].sum().reset_index()
            fig = px.pie(
                svc, names="Service_Type", values="Amount",
                title="🩺 Revenue by Service Type",
                hole=0.4,
            )
            st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        if "Status" in master.columns:
            st_df = master["Status"].value_counts().reset_index()
            st_df.columns = ["Status", "Count"]
            fig = px.bar(
                st_df, x="Status", y="Count",
                title="📋 Payment Status",
                color="Status",
            )
            st.plotly_chart(fig, use_container_width=True)

    with col4:
        if "Doctor" in master.columns and "Amount" in master.columns:
            top_docs = (
                master.groupby("Doctor")["Amount"].sum()
                .nlargest(10).reset_index()
            )
            fig = px.bar(
                top_docs, x="Amount", y="Doctor", orientation="h",
                title="🏆 Top 10 Doctors by Revenue",
            )
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

# ----- Tab 2: By Clinic -----
with tab2:
    if "Location" in master.columns:
        clinic_stats = master.groupby("Location").agg(
            Total_Amount=("Amount", "sum"),
            Patients=("Patient_Name", "nunique") if "Patient_Name" in master
                     else ("Amount", "count"),
            Avg_Sale=("Amount", "mean"),
        ).round(2).reset_index()

        st.dataframe(
            clinic_stats.style.format({
                "Total_Amount": "${:,.0f}",
                "Avg_Sale": "${:,.2f}",
            }),
            use_container_width=True,
        )

        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(
                clinic_stats, x="Location", y="Total_Amount",
                title="Revenue by Clinic", color="Location",
            )
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            fig = px.bar(
                clinic_stats, x="Location", y="Avg_Sale",
                title="Avg Sale by Clinic", color="Location",
            )
            st.plotly_chart(fig, use_container_width=True)

# ----- Tab 3: By Doctor -----
with tab3:
    if "Doctor" in master.columns:
        doc_stats = master.groupby("Doctor").agg(
            Patients=("Patient_Name", "nunique") if "Patient_Name" in master
                     else ("Amount", "count"),
            Revenue=("Amount", "sum"),
            Avg_Sale=("Amount", "mean"),
        ).round(2).sort_values("Revenue", ascending=False).reset_index()

        st.dataframe(
            doc_stats.style.format({
                "Revenue": "${:,.0f}",
                "Avg_Sale": "${:,.2f}",
            }),
            use_container_width=True,
        )

# ----- Tab 4: Raw Data -----
with tab4:
    st.dataframe(master, use_container_width=True)

    csv = master.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download Filtered Data (CSV)",
        csv,
        "healthcare_filtered.csv",
        "text/csv",
    )

st.caption(
    f"📁 Source: output/Healthcare_Data_Automation.xlsx · "
    f"{len(master):,} rows shown"
)
