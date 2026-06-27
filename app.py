import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(
    page_title="Annual Budget Dashboard",
    page_icon="📊",
    layout="wide"
)

# ==========================
# HEADER
# ==========================
st.markdown(
    """
    <div style='text-align:center'>
        <h1>USMAN PUBLIC SCHOOL SYSTEM</h1>
        <h3>Campus 45</h3>
        <h2>Annual Budget Dashboard</h2>
        <h3>FY 2026-27</h3>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("---")

# ==========================
# FILE UPLOADS
# ==========================
col1, col2 = st.columns(2)

with col1:
    income_file = st.file_uploader(
        "Upload Income Budget File",
        type=["xlsx"]
    )

with col2:
    expense_file = st.file_uploader(
        "Upload Expense Budget File",
        type=["xlsx"]
    )

if income_file and expense_file:

    income = pd.read_excel(
        income_file,
        sheet_name="Budget FY 2026-27"
    )

    expense = pd.read_excel(
        expense_file,
        sheet_name="Budget FY 2026-27"
    )

    months = [
        'Jul-26','Aug-26','Sep-26','Oct-26',
        'Nov-26','Dec-26','Jan-27','Feb-27',
        'Mar-27','Apr-27','May-27','Jun-27'
    ]

    income_totals = income[months].sum()
    expense_totals = expense[months].sum()
    surplus = income_totals - expense_totals

    total_income = income_totals.sum()
    total_expense = expense_totals.sum()
    total_surplus = surplus.sum()

    surplus_margin = (
        total_surplus / total_income * 100
        if total_income != 0 else 0
    )

    # ==========================
    # KPI CARDS
    # ==========================
    st.subheader("Executive Summary")

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Total Budgeted Income",
        f"{total_income:,.0f}"
    )

    c2.metric(
        "Total Budgeted Expenses",
        f"{total_expense:,.0f}"
    )

    c3.metric(
        "Budgeted Surplus",
        f"{total_surplus:,.0f}"
    )

    c4.metric(
        "Surplus Margin %",
        f"{surplus_margin:.2f}%"
    )

    st.markdown("---")

    # ==========================
    # MONTHLY DATA
    # ==========================
    pnl = pd.DataFrame({
        'Month': months,
        'Income': income_totals.values,
        'Expenses': expense_totals.values,
        'Surplus': surplus.values
    })

    # ==========================
    # CHARTS
    # ==========================
    c1, c2 = st.columns(2)

    with c1:

        fig = px.bar(
            pnl,
            x='Month',
            y=['Income', 'Expenses'],
            barmode='group',
            title='Monthly Income vs Expenses'
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    with c2:

        fig2 = px.line(
            pnl,
            x='Month',
            y='Surplus',
            markers=True,
            title='Monthly Surplus Trend'
        )

        st.plotly_chart(
            fig2,
            use_container_width=True
        )

    # ==========================
    # TOP 10 EXPENSE HEADS
    # ==========================
    expense['Total'] = expense[months].sum(axis=1)

    top10 = (
        expense[['Account', 'Total']]
        .sort_values(
            'Total',
            ascending=False
        )
        .head(10)
    )

    fig3 = px.bar(
        top10,
        x='Total',
        y='Account',
        orientation='h',
        title='Top 10 Expense Heads'
    )

    st.plotly_chart(
        fig3,
        use_container_width=True
    )

    # ==========================
    # MONTHLY P&L TABLE
    # ==========================
    st.subheader(
        "Monthly Budgeted Profit & Loss Statement"
    )

    pnl_display = pnl.copy()

    pnl_display['Income'] = (
        pnl_display['Income']
        .map('{:,.0f}'.format)
    )

    pnl_display['Expenses'] = (
        pnl_display['Expenses']
        .map('{:,.0f}'.format)
    )

    pnl_display['Surplus'] = (
        pnl_display['Surplus']
        .map('{:,.0f}'.format)
    )

    st.dataframe(
        pnl_display,
        use_container_width=True
    )

    # ==========================
    # DOWNLOAD P&L
    # ==========================
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine='xlsxwriter'
    ) as writer:

        pnl.to_excel(
            writer,
            sheet_name='Budgeted P&L',
            index=False
        )

    st.download_button(
        label="📥 Download Budgeted P&L",
        data=output.getvalue(),
        file_name='Budgeted_PnL.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
