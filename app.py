import streamlit as st
import pandas as pd
import numpy as np
import calendar
import plotly.express as px
from st_aggrid import AgGrid
from io import BytesIO

# ---------------- Page Config ----------------
st.set_page_config(page_title="My Request Astras Source Reporting Dashboard", layout="wide")

st.title("📊 My Request Astras Source Reporting Dashboard")

# ---------------- File Upload ----------------
st.markdown("### 📂 Upload Input Files")

col1, col2 = st.columns(2)
with col1:
    export_file = st.file_uploader("Upload the Sourcing Events file (English version)", type=["xlsx"], key="export")
with col2:
    sms_file = st.file_uploader("Upload Supplier Market Structure File (English version)", type=["xlsx"], key="sms")

# ---------------- Fallback Message ----------------
if not export_file or not sms_file:
    st.warning("📂 Please upload both Excel files to proceed.")
    st.stop()

# ---------------- Data Processing ----------------
df1 = pd.read_excel(export_file, sheet_name=0)
df1["Purchasing Category (PC)"] = df1["Material"].str.extract(r"\[(\d+)\]")
df1["Purchasing Category (PC)"] = df1["Purchasing Category (PC)"].astype(float).astype("Int64")
df1 = df1[~df1["Name"].str.lower().str.contains("test", na=False)]

df2 = pd.read_excel(sms_file, sheet_name="New Structure")
df2 = df2[["Purchasing Category (PC)", "Supply Market Cluster (SMC)"]].drop_duplicates()

df_merged = df1.merge(df2, on="Purchasing Category (PC)", how="left")
df_merged["Date of entry"] = pd.to_datetime(df_merged["Date of entry"], format="%d/%m/%Y", errors="coerce")
df_merged.rename(columns={"Supply Market Cluster (SMC)": "Groups"}, inplace=True)

keep_values = [
    "PHOENIX CONTACT E-Mobility GmbH",
    "Phoenix Contact GmbH & Co. KG - Werkzeugbau",
]
df_merged["Organization1"] = np.where(
    df_merged["Organization"].isin(keep_values),
    df_merged["Organization"],
    "Phoenix Contact - GPN",
)

df_merged["month"] = df_merged["Date of entry"].dt.month
df_merged["month_name"] = df_merged["month"].apply(lambda x: calendar.month_name[x] if pd.notna(x) else "")
month_order = list(calendar.month_name[1:])
df_merged["month_name"] = pd.Categorical(df_merged["month_name"], categories=month_order, ordered=True)

# ---------------- Filters ----------------
st.markdown("### 🎛️ Filters")

col1, col2 = st.columns(2)
with col1:
    orgs = st.multiselect("Select Organizations", options=df_merged["Organization1"].unique(), default=df_merged["Organization1"].unique())
with col2:
    months = st.multiselect("Select Months", options=month_order, default=month_order)

filtered_df = df_merged[
    (df_merged["Organization1"].isin(orgs)) & (df_merged["month_name"].isin(months))
]

# ---------------- Total Events ----------------
total_events = len(filtered_df)
st.markdown(f"### 🔢 Number of Events: **{total_events}**")

# ---------------- Visualizations ----------------
st.header("📈 Visualizations")

monthly_counts = filtered_df.groupby(["Organization1", "month_name"]).size().reset_index(name="count")
monthly_counts["month_name"] = pd.Categorical(monthly_counts["month_name"], categories=month_order, ordered=True)

fig1 = px.bar(
    monthly_counts,
    x="month_name",
    y="count",
    color="Organization1",
    barmode="group",
    text="count",
    category_orders={"month_name": month_order},
    title="Monthly Distribution per Organization",
)
fig1.update_traces(textposition="outside")
st.plotly_chart(fig1, use_container_width=True)

df_gpn = filtered_df[filtered_df["Organization1"] == "Phoenix Contact - GPN"].copy()
if not df_gpn.empty:
    pivot_cluster = df_gpn.pivot_table(index="Groups", columns="month_name", aggfunc="size", fill_value=0)
    pivot_cluster = pivot_cluster.reindex(columns=month_order, fill_value=0)

    fig2 = px.bar(
        pivot_cluster,
        x=pivot_cluster.columns,
        y=pivot_cluster.index,
        orientation="h",
        barmode="stack",
        category_orders={"month_name": month_order},
        title="Event Distribution per Cluster by Month (GPN)",
    )
    st.plotly_chart(fig2, use_container_width=True)

    pivot_org = df_gpn.pivot_table(index="Organization", columns="month_name", aggfunc="size", fill_value=0)
    pivot_org = pivot_org.reindex(columns=month_order, fill_value=0)

    fig3 = px.bar(
        pivot_org,
        x=pivot_org.columns,
        y=pivot_org.index,
        orientation="h",
        barmode="stack",
        category_orders={"month_name": month_order},
        title="Event Distribution per Organization by Month (GPN)",
    )
    st.plotly_chart(fig3, use_container_width=True)

# ---------------- Tables ----------------
st.header("📑 Detailed Data Tables")

def make_downloadable_excel(df, filename="table.xlsx"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buffer.getvalue()

st.subheader("Merged Export Data")
AgGrid(filtered_df, height=300)
st.download_button("⬇️ Download Export Data", make_downloadable_excel(filtered_df), "export_data.xlsx")

pivot_table1 = filtered_df.pivot_table(index="month_name", columns="Organization1", aggfunc="size", fill_value=0).reindex(month_order).reset_index()
st.subheader("Pivot Table: Month vs Organization")
AgGrid(pivot_table1, height=300)
st.download_button("⬇️ Download Pivot Table 1", make_downloadable_excel(pivot_table1), "pivot_table1.xlsx")

if not df_gpn.empty:
    pivot_table2 = df_gpn.pivot_table(index="month_name", columns="Groups", aggfunc="size", fill_value=0).reindex(month_order).reset_index()
    st.subheader("Pivot Table: Month vs Groups (GPN only)")
    AgGrid(pivot_table2, height=300)
    st.download_button("⬇️ Download Pivot Table 2", make_downloadable_excel(pivot_table2), "pivot_table2.xlsx")
