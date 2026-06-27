import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(
    page_title="Annual Budget Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================
# HEADER
# ==========================
st.markdown(
    """
    <div style='text-align:center;
                padding:25px;
                border-radius:20px;
                background-color:#1E1E1E;
                margin-bottom:20px;'>

        <h1 style='color:white;'>
        USMAN PUBLIC SCHOOL SYSTEM
        </h1>

        <h3 style='color:#B0B0B0;'>
        Campus 45
        </h3>

        <h2 style='color:#2ECC71;'>
        Annual Budget Dashboard
        </h2>

        <h3 style='color:#B0B0B0;'>
        FY 2026-27
        </h3>

    </div>
    """,
    unsafe_allow_html=True
)

# ==========================
# SIDEBAR
# ==========================
with st.sidebar:

    st.image(
        "https://img.icons8.com/fluency/96/dashboard-layout.png",
        width=80
    )

    st.markdown("## 📊 Budget Dashboard")

    st.markdown(
        """
        Upload:

        ✅ Income Budget File

        ✅ Expense Budget File
        """
    )

    income_file = st.file_uploader(
        "💚 Income Budget File",
        type="xlsx"
    )

    expense_file = st.file_uploader(
        "❤️ Expense Budget File",
        type="xlsx"
    )

    st.markdown("---")

    st.info(
        """
        Upload both files to generate:

        • Executive Dashboard

        • Profit & Loss Statement

        • Downloadable Excel P&L
        """
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
    "💚 Total Income",
    f"{total_income:,.0f}"
)

c2.metric(
    "❤️ Total Expenses",
    f"{total_expense:,.0f}"
)

c3.metric(
    "💙 Budgeted Surplus",
    f"{total_surplus:,.0f}"
)

c4.metric(
    "🟨 Surplus Margin %",
    f"{surplus_margin:.2f}%"
)
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
    color_discrete_sequence=[
        '#2ECC71',
        '#E74C3C'
    ],
    title='Monthly Income vs Expenses'
)

fig.update_layout(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    font=dict(size=14),
    height=500
)

    with c2:

       fig2 = px.line(
    pnl,
    x='Month',
    y='Surplus',
    markers=True,
    title='Monthly Surplus Trend'
)

fig2.update_traces(
    line=dict(
        color='#3498DB',
        width=5
    )
)

fig2.update_layout(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    height=500
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
    color='Total',
    color_continuous_scale='Reds',
    title='Top 10 Expense Heads'
)

fig3.update_layout(
    height=550,
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)'
)

    st.plotly_chart(
        fig3,
        use_container_width=True
    )

    # ==========================
    # MONTHLY P&L TABLE
    # ==========================
   st.subheader(
    "📋 Monthly Budgeted Profit & Loss Statement"
)

st.dataframe(
    pnl.style.format({
        'Income':'{:,.0f}',
        'Expenses':'{:,.0f}',
        'Surplus':'{:,.0f}'
    }),
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
        startrow=3,
        index=False
    )

    workbook = writer.book
    worksheet = writer.sheets['Budgeted P&L']

    title = workbook.add_format({
        'bold':True,
        'font_size':18,
        'bg_color':'#1F4E78',
        'font_color':'white',
        'align':'center',
        'valign':'vcenter'
    })

    header = workbook.add_format({
        'bold':True,
        'bg_color':'#D9EAD3',
        'border':1
    })

    money = workbook.add_format({
        'num_format':'#,##0',
        'border':1
    })

    total_format = workbook.add_format({
        'bold':True,
        'bg_color':'#BDD7EE',
        'num_format':'#,##0',
        'border':1
    })

    worksheet.merge_range(
        'A1:D1',
        'USMAN PUBLIC SCHOOL SYSTEM - Campus 45',
        title
    )

    worksheet.merge_range(
        'A2:D2',
        'Annual Budget FY 2026-27',
        title
    )

    for col_num, value in enumerate(pnl.columns.values):
        worksheet.write(
            3,
            col_num,
            value,
            header
        )

    worksheet.set_column('A:A',20)
    worksheet.set_column('B:D',22)

    total_row = len(pnl) + 4

    worksheet.write(
        total_row,
        0,
        'Annual Total',
        total_format
    )

    worksheet.write(
        total_row,
        1,
        pnl['Income'].sum(),
        total_format
    )

    worksheet.write(
        total_row,
        2,
        pnl['Expenses'].sum(),
        total_format
    )

    worksheet.write(
        total_row,
        3,
        pnl['Surplus'].sum(),
        total_format
    )
