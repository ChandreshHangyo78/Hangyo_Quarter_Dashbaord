import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from PIL import Image

# --- Custom Indian Number Formatter ---
def format_in_indian_style(num):
    num = int(num)
    s = str(num)
    if len(s) <= 3:
        return s
    else:
        last3 = s[-3:]
        remaining = s[:-3]
        parts = []
        while len(remaining) > 2:
            parts.insert(0, remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            parts.insert(0, remaining)
        return ','.join(parts) + ',' + last3

# --- Display Company Logo ---
logo = Image.open("hangyo_logo.png")
st.image(logo, width=180)

# --- Load Excel Files ---
main_df = pd.read_excel("Sales_Expense_Dashbaord_Excel.xlsx")
expense_df = pd.read_excel("Dashbaord_Approach_Excel_File.xlsx")

# --- Clean column names ---
main_df.columns = main_df.columns.str.strip()
expense_df.columns = expense_df.columns.str.strip()

# --- Sidebar Filters ---
quarters = ["Q1", "Q2", "Q3", "Q4"]
selected_quarter = st.sidebar.selectbox("Select Quarter", quarters)

regions = ["All"] + sorted(main_df["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", regions)

if selected_region != "All":
    hubs = ["All"] + sorted(main_df[main_df["Region"] == selected_region]["Hub"].dropna().unique())
else:
    hubs = ["All"] + sorted(main_df["Hub"].dropna().unique())
selected_hub = st.sidebar.selectbox("Select Hub", hubs)

expense_types = ["All"] + sorted(expense_df["Expense Type"].dropna().unique())
selected_expense = st.sidebar.selectbox("Select Expense Type", expense_types)

# --- Filter Main Summary Data ---
filtered_main = main_df.copy()
if selected_region != "All":
    filtered_main = filtered_main[filtered_main["Region"] == selected_region]
if selected_hub != "All":
    filtered_main = filtered_main[filtered_main["Hub"] == selected_hub]

# --- KPI Section ---
st.title("📊 Sales & Expense Dashboard")
q_exp = f"{selected_quarter}_EXPENSES"
q_sales = f"{selected_quarter}_SALES"

if not filtered_main.empty:
    total_exp = filtered_main[q_exp].sum()
    total_sales = filtered_main[q_sales].sum()
    exp_pct = (total_exp / total_sales * 100) if total_sales > 0 else 0

    st.metric(f"{selected_quarter} Expenses", f"₹ {format_in_indian_style(total_exp)}")
    st.metric(f"{selected_quarter} Sales", f"₹ {format_in_indian_style(total_sales)}")
    st.metric(f"{selected_quarter} Expense % of Sales", f"{exp_pct:.2f}%")

# --- Filter Expense-Type Breakdown Data ---
filtered_exp = expense_df.copy()
if selected_region != "All":
    filtered_exp = filtered_exp[filtered_exp["Region"] == selected_region]
if selected_hub != "All":
    filtered_exp = filtered_exp[filtered_exp["Hub"] == selected_hub]
if selected_expense != "All":
    filtered_exp = filtered_exp[filtered_exp["Expense Type"] == selected_expense]

# --- Format Claim Columns as Percentages ---
claim_cols = [f"{q}_Claim" for q in ["Q1", "Q2", "Q3", "Q4"] if f"{q}_Claim" in filtered_exp.columns]
for col in claim_cols:
    filtered_exp[col] = (filtered_exp[col] * 100).round(2).astype(str) + '%'

# --- Pie Chart ---
st.subheader("🧾 Expense Type Contribution")
if not filtered_exp.empty and q_exp in filtered_exp.columns:
    pie_fig = px.pie(filtered_exp, names="Expense Type", values=q_exp)
    st.plotly_chart(pie_fig)

# --- Bar Chart: Expense vs Sales vs Expense % ---
st.subheader("📈 Quarter-wise Expense vs Sales")
if not filtered_main.empty:
    quarters_all = ["Q1", "Q2", "Q3", "Q4"]
    bar_data = pd.DataFrame({
        "Quarter": quarters_all,
        "Expense": [filtered_main[f"{q}_EXPENSES"].sum() for q in quarters_all],
        "Sales": [filtered_main[f"{q}_SALES"].sum() for q in quarters_all],
    })
    bar_data["Expense %"] = (bar_data["Expense"] / bar_data["Sales"]) * 100

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=bar_data["Quarter"],
        y=bar_data["Expense"],
        name="Expenses",
        text=[f"₹ {format_in_indian_style(x)}" for x in bar_data["Expense"]],
        textposition="outside"
    ))

    fig.add_trace(go.Bar(
        x=bar_data["Quarter"],
        y=bar_data["Sales"],
        name="Sales",
        text=[f"₹ {format_in_indian_style(x)}" for x in bar_data["Sales"]],
        textposition="outside"
    ))

    fig.add_trace(go.Scatter(
        x=bar_data["Quarter"],
        y=bar_data["Expense %"],
        mode='lines+markers+text',
        name="Expense %",
        text=[f"{x:.2f}%" for x in bar_data["Expense %"]],
        textposition="top center",
        yaxis="y2"
    ))

    fig.update_layout(
        yaxis=dict(title="Amount (₹)", tickformat=","),
        yaxis2=dict(title="Expense %", overlaying="y", side="right", showgrid=False),
        barmode='group',
        title="Expenses, Sales and % Comparison by Quarter"
    )

    st.plotly_chart(fig)

# --- Show Data Table ---
st.subheader("📋 Filtered Detailed Expense Type Data")
st.dataframe(filtered_exp)

# --- Summary Insights ---
st.markdown("---")
st.subheader("📌 Summary Insights")

if not filtered_exp.empty:
    top_expense = filtered_exp.sort_values(by=q_exp, ascending=False).iloc[0]
    claim_percentages = []
    for col in claim_cols:
        try:
            val = float(top_expense[col].replace('%', ''))
            claim_percentages.append(val)
        except:
            continue
    avg_claim_pct = sum(claim_percentages) / len(claim_percentages) if claim_percentages else 0

    st.markdown(f"""
    - **Top Expense Type** in {selected_quarter}: `{top_expense['Expense Type']}` → ₹ {format_in_indian_style(top_expense[q_exp])}
    - **Average Claim % across quarters**: `{avg_claim_pct:.2f}%`
    - **Total Expense Entries Analyzed**: `{len(filtered_exp)}`
    """)
else:
    st.info("No data available for insights.")
