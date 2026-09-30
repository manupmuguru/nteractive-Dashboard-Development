# nteractive-Dashboard-Development
# ================================================================
#        INTERACTIVE BUSINESS INTELLIGENCE DASHBOARD
# ================================================================
#
# Project:
# Interactive Dashboard Development
#
# Technology:
# Python + Streamlit + Pandas + Plotly
#
# Dashboard Features:
# 1. Upload CSV / Excel dataset
# 2. Automatic column detection
# 3. KPI cards
# 4. Date range filter
# 5. Region filter
# 6. Category filter
# 7. Revenue trend
# 8. Revenue by region
# 9. Revenue by category
# 10. Active users analysis
# 11. Churn percentage
# 12. Average ticket size
# 13. Monthly analysis
# 14. Interactive charts
# 15. Data table
# 16. Download filtered data
#
# Run:
# streamlit run app.py
#
# ================================================================


# ================================================================
# 1. IMPORT LIBRARIES
# ================================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import os


# ================================================================
# 2. PAGE CONFIGURATION
# ================================================================

st.set_page_config(
    page_title="Business Intelligence Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ================================================================
# 3. CUSTOM CSS
# ================================================================

st.markdown(
    """
    <style>

    /* Main page */

    .main {
        padding-top: 1rem;
    }


    /* KPI cards */

    .kpi-card {
        background: white;
        padding: 20px;
        border-radius: 15px;
        border: 1px solid #e6e6e6;
        box-shadow: 0px 3px 12px rgba(0,0,0,0.08);
        min-height: 130px;
    }


    .kpi-title {
        font-size: 15px;
        color: #666666;
        font-weight: 600;
        margin-bottom: 8px;
    }


    .kpi-value {
        font-size: 30px;
        font-weight: 700;
        color: #222222;
    }


    .kpi-subtitle {
        font-size: 13px;
        color: #777777;
        margin-top: 5px;
    }


    /* Dashboard title */

    .dashboard-title {
        font-size: 38px;
        font-weight: 800;
        color: #202124;
        margin-bottom: 5px;
    }


    .dashboard-subtitle {
        font-size: 17px;
        color: #666666;
        margin-bottom: 25px;
    }


    /* Section heading */

    .section-title {
        font-size: 22px;
        font-weight: 700;
        margin-top: 20px;
        margin-bottom: 10px;
    }


    /* Info boxes */

    .info-box {
        padding: 15px;
        border-radius: 12px;
        background: #f4f7ff;
        border-left: 5px solid #6366f1;
        margin-bottom: 15px;
    }


    /* Footer */

    .footer {
        text-align: center;
        padding: 25px;
        color: #777777;
        font-size: 14px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ================================================================
# 4. HELPER FUNCTIONS
# ================================================================

def find_column(df, possible_names):

    """
    Automatically find a column from a list of possible names.
    """

    columns = {
        str(col).strip().lower().replace(" ", "_"): col
        for col in df.columns
    }

    for name in possible_names:

        key = (
            name
            .strip()
            .lower()
            .replace(" ", "_")
        )

        if key in columns:
            return columns[key]

    # Partial matching

    for normalized, original in columns.items():

        for name in possible_names:

            name = (
                name
                .strip()
                .lower()
                .replace(" ", "_")
            )

            if name in normalized:
                return original

    return None


# ================================================================
# 5. LOAD DATA FUNCTION
# ================================================================

def load_uploaded_file(uploaded_file):

    """
    Load CSV or Excel file.
    """

    if uploaded_file is None:
        return None

    file_name = uploaded_file.name.lower()

    try:

        if file_name.endswith(".csv"):

            data = pd.read_csv(
                uploaded_file
            )

        elif (
            file_name.endswith(".xlsx")
            or file_name.endswith(".xls")
        ):

            data = pd.read_excel(
                uploaded_file
            )

        else:

            st.error(
                "Please upload a CSV or Excel file."
            )

            return None

        return data

    except Exception as error:

        st.error(
            f"Unable to read file: {error}"
        )

        return None


# ================================================================
# 6. CLEAN DATA
# ================================================================

def clean_data(data):

    df = data.copy()

    # Remove duplicate rows

    df = df.drop_duplicates()

    # Clean column names

    df.columns = [
        str(column)
        .strip()
        .replace(" ", "_")
        .replace("-", "_")
        .lower()
        for column in df.columns
    ]

    # Strip text values

    for column in df.select_dtypes(
        include="object"
    ).columns:

        df[column] = (
            df[column]
            .astype(str)
            .str.strip()
        )

    return df


# ================================================================
# 7. AUTOMATIC COLUMN DETECTION
# ================================================================

def detect_columns(df):

    detected = {}

    detected["date"] = find_column(
        df,
        [
            "date",
            "order_date",
            "transaction_date",
            "sale_date",
            "purchase_date"
        ]
    )

    detected["revenue"] = find_column(
        df,
        [
            "revenue",
            "sales",
            "sales_amount",
            "total_sales",
            "total_revenue",
            "amount",
            "price"
        ]
    )

    detected["user"] = find_column(
        df,
        [
            "user",
            "user_id",
            "customer",
            "customer_id",
            "client_id",
            "userid"
        ]
    )

    detected["region"] = find_column(
        df,
        [
            "region",
            "state",
            "area",
            "territory",
            "location"
        ]
    )

    detected["category"] = find_column(
        df,
        [
            "category",
            "product_category",
            "product",
            "segment",
            "type"
        ]
    )

    detected["churn"] = find_column(
        df,
        [
            "churn",
            "churned",
            "is_churned",
            "churn_status"
        ]
    )

    detected["quantity"] = find_column(
        df,
        [
            "quantity",
            "qty",
            "units",
            "items"
        ]
    )

    return detected


# ================================================================
# 8. PREPARE DATA
# ================================================================

def prepare_data(df, detected):

    data = df.copy()

    # ------------------------------------------------------------
    # DATE
    # ------------------------------------------------------------

    if detected["date"] is not None:

        data[detected["date"]] = pd.to_datetime(
            data[detected["date"]],
            errors="coerce"
        )

    # ------------------------------------------------------------
    # REVENUE
    # ------------------------------------------------------------

    if detected["revenue"] is not None:

        revenue_column = detected["revenue"]

        data[revenue_column] = (
            data[revenue_column]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("₹", "", regex=False)
            .str.replace("$", "", regex=False)
            .str.replace("€", "", regex=False)
        )

        data[revenue_column] = pd.to_numeric(
            data[revenue_column],
            errors="coerce"
        )

        data[revenue_column] = (
            data[revenue_column]
            .fillna(0)
        )

    # ------------------------------------------------------------
    # CHURN
    # ------------------------------------------------------------

    if detected["churn"] is not None:

        churn_column = detected["churn"]

        data[churn_column] = (
            data[churn_column]
            .astype(str)
            .str.lower()
            .str.strip()
        )

    return data


# ================================================================
# 9. CALCULATE CHURN RATE
# ================================================================

def calculate_churn_rate(data, churn_column):

    if churn_column is None:
        return 0.0

    churn_values = (
        data[churn_column]
        .astype(str)
        .str.lower()
        .str.strip()
    )

    churned_values = [
        "yes",
        "y",
        "true",
        "1",
        "churned",
        "churn"
    ]

    churned = churn_values.isin(
        churned_values
    ).sum()

    total = len(data)

    if total == 0:
        return 0.0

    return (
        churned /
        total *
        100
    )


# ================================================================
# 10. ACTIVE USER CALCULATION
# ================================================================

def calculate_active_users(data, user_column):

    if user_column is None:

        return len(data)

    return data[user_column].nunique()


# ================================================================
# 11. AVERAGE TICKET SIZE
# ================================================================

def calculate_average_ticket(
    data,
    revenue_column
):

    if revenue_column is None:
        return 0.0

    if len(data) == 0:
        return 0.0

    return (
        data[revenue_column].sum()
        /
        len(data)
    )


# ================================================================
# 12. APPLICATION HEADER
# ================================================================

st.markdown(
    """
    <div class="dashboard-title">
        📊 Interactive Business Intelligence Dashboard
    </div>

    <div class="dashboard-subtitle">
        Dynamic analysis of revenue, users, churn and regional performance
    </div>
    """,
    unsafe_allow_html=True
)


# ================================================================
# 13. SIDEBAR
# ================================================================

st.sidebar.title(
    "⚙️ Dashboard Controls"
)

st.sidebar.markdown(
    """
    Upload your structured business dataset
    and use the filters to explore the data.
    """
)


# ================================================================
# 14. FILE UPLOADER
# ================================================================

uploaded_file = st.sidebar.file_uploader(
    "📁 Upload Dataset",
    type=[
        "csv",
        "xlsx",
        "xls"
    ]
)


# ================================================================
# 15. SAMPLE DATA GENERATOR
# ================================================================

use_sample_data = st.sidebar.checkbox(
    "Use sample dataset",
    value=False
)


# ================================================================
# 16. CREATE SAMPLE DATASET
# ================================================================

def create_sample_data():

    np.random.seed(42)

    number_of_records = 1000

    dates = pd.date_range(
        start="2025-01-01",
        end="2026-08-31",
        periods=number_of_records
    )

    regions = [
        "Bengaluru",
        "Chennai",
        "Mumbai",
        "Delhi",
        "Hyderabad",
        "Pune",
        "Kolkata",
        "Mysuru"
    ]

    categories = [
        "Electronics",
        "Software",
        "Services",
        "Accessories",
        "Hardware"
    ]

    users = [
        f"USER{number:04d}"
        for number in np.random.randint(
            1,
            500,
            number_of_records
        )
    ]

    revenue = np.random.gamma(
        shape=4,
        scale=1500,
        size=number_of_records
    )

    churn = np.random.choice(
        [
            "Yes",
            "No"
        ],
        size=number_of_records,
        p=[
            0.15,
            0.85
        ]
    )

    sample = pd.DataFrame({

        "date": dates,

        "revenue": revenue,

        "user_id": users,

        "region": np.random.choice(
            regions,
            number_of_records
        ),

        "category": np.random.choice(
            categories,
            number_of_records
        ),

        "churn": churn

    })

    return sample


# ================================================================
# 17. SELECT DATA SOURCE
# ================================================================

if use_sample_data:

    df = create_sample_data()

    st.sidebar.success(
        "Sample dataset loaded."
    )

elif uploaded_file is not None:

    df = load_uploaded_file(
        uploaded_file
    )

else:

    df = None


# ================================================================
# 18. SHOW INSTRUCTIONS IF NO DATA
# ================================================================

if df is None:

    st.info(
        """
        👋 Welcome!

        Upload a CSV/Excel dataset from the sidebar,
        or select **Use sample dataset** to test the dashboard.
        """
    )

    st.markdown(
        """
        ### Dashboard Includes

        - 💰 Total Revenue
        - 👥 Active Users
        - 📉 Churn %
        - 🎫 Average Ticket Size
        - 📈 Revenue Trend
        - 🌎 Regional Performance
        - 🏷️ Category Analysis
        - 📊 Churn Analysis
        - 🔎 Interactive Filters
        - 📥 Download Filtered Data
        """
    )

    st.stop()


# ================================================================
# 19. CLEAN DATA
# ================================================================

df = clean_data(df)


# ================================================================
# 20. DETECT COLUMNS
# ================================================================

detected = detect_columns(df)


# ================================================================
# 21. PREPARE DATA
# ================================================================

df = prepare_data(
    df,
    detected
)


# ================================================================
# 22. SHOW DATASET INFORMATION
# ================================================================

st.sidebar.markdown(
    "---"
)

st.sidebar.subheader(
    "🔍 Detected Columns"
)


for key, value in detected.items():

    if value is not None:

        st.sidebar.write(
            f"**{key.title()}:** {value}"
        )

    else:

        st.sidebar.write(
            f"**{key.title()}:** Not found"
        )


# ================================================================
# 23. SIDEBAR FILTERS
# ================================================================

st.sidebar.markdown(
    "---"
)

st.sidebar.subheader(
    "🎛️ Filters"
)


# ================================================================
# DATE FILTER
# ================================================================

filtered_df = df.copy()


if detected["date"] is not None:

    date_column = detected["date"]

    valid_dates = filtered_df[
        date_column
    ].dropna()


    if len(valid_dates) > 0:

        minimum_date = valid_dates.min().date()

        maximum_date = valid_dates.max().date()


        date_range = st.sidebar.date_input(
            "📅 Date Range",
            value=(
                minimum_date,
                maximum_date
            ),
            min_value=minimum_date,
            max_value=maximum_date
        )


        if isinstance(
            date_range,
            tuple
        ) and len(date_range) == 2:

            start_date = pd.Timestamp(
                date_range[0]
            )

            end_date = (
                pd.Timestamp(
                    date_range[1]
                )
                +
                pd.Timedelta(
                    days=1
                )
            )


            filtered_df = filtered_df[
                (
                    filtered_df[
                        date_column
                    ] >= start_date
                )
                &
                (
                    filtered_df[
                        date_column
                    ] < end_date
                )
            ]


# ================================================================
# REGION FILTER
# ================================================================

if detected["region"] is not None:

    region_column = detected["region"]

    regions = sorted(
        filtered_df[
            region_column
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )


    selected_regions = st.sidebar.multiselect(
        "🌍 Region",
        options=regions,
        default=regions
    )


    if len(selected_regions) > 0:

        filtered_df = filtered_df[
            filtered_df[
                region_column
            ].astype(str).isin(
                selected_regions
            )
        ]


# ================================================================
# CATEGORY FILTER
# ================================================================

if detected["category"] is not None:

    category_column = detected["category"]

    categories = sorted(
        filtered_df[
            category_column
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )


    selected_categories = st.sidebar.multiselect(
        "🏷️ Category",
        options=categories,
        default=categories
    )


    if len(selected_categories) > 0:

        filtered_df = filtered_df[
            filtered_df[
                category_column
            ].astype(str).isin(
                selected_categories
            )
        ]


# ================================================================
# 24. RESET / FILTER INFORMATION
# ================================================================

st.sidebar.markdown(
    "---"
)

st.sidebar.metric(
    "Filtered Records",
    len(filtered_df)
)


# ================================================================
# 25. KPI CALCULATIONS
# ================================================================

if detected["revenue"] is not None:

    total_revenue = filtered_df[
        detected["revenue"]
    ].sum()

else:

    total_revenue = 0


active_users = calculate_active_users(
    filtered_df,
    detected["user"]
)


churn_rate = calculate_churn_rate(
    filtered_df,
    detected["churn"]
)


average_ticket = calculate_average_ticket(
    filtered_df,
    detected["revenue"]
)


# ================================================================
# 26. KPI CARDS
# ================================================================

st.markdown(
    '<div class="section-title">📌 Key Performance Indicators</div>',
    unsafe_allow_html=True
)


kpi1, kpi2, kpi3, kpi4 = st.columns(4)


with kpi1:

    st.markdown(
        f"""
        <div class="kpi-card">

            <div class="kpi-title">
                💰 Total Revenue
            </div>

            <div class="kpi-value">
                ₹{total_revenue:,.2f}
            </div>

            <div class="kpi-subtitle">
                Revenue from filtered records
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


with kpi2:

    st.markdown(
        f"""
        <div class="kpi-card">

            <div class="kpi-title">
                👥 Active Users
            </div>

            <div class="kpi-value">
                {active_users:,}
            </div>

            <div class="kpi-subtitle">
                Unique customers/users
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


with kpi3:

    st.markdown(
        f"""
        <div class="kpi-card">

            <div class="kpi-title">
                📉 Churn %
            </div>

            <div class="kpi-value">
                {churn_rate:.2f}%
            </div>

            <div class="kpi-subtitle">
                Customer churn rate
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


with kpi4:

    st.markdown(
        f"""
        <div class="kpi-card">

            <div class="kpi-title">
                🎫 Average Ticket Size
            </div>

            <div class="kpi-value">
                ₹{average_ticket:,.2f}
            </div>

            <div class="kpi-subtitle">
                Average revenue per transaction
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


# ================================================================
# 27. DATASET SUMMARY
# ================================================================

st.markdown(
    '<div class="section-title">📋 Dataset Summary</div>',
    unsafe_allow_html=True
)


summary_col1, summary_col2, summary_col3 = st.columns(3)


with summary_col1:

    st.info(
        f"**Total Records:** {len(filtered_df):,}"
    )


with summary_col2:

    st.info(
        f"**Columns:** {len(filtered_df.columns):,}"
    )


with summary_col3:

    st.info(
        f"**Original Records:** {len(df):,}"
    )


# ================================================================
# 28. REVENUE TREND
# ================================================================

st.markdown(
    '<div class="section-title">📈 Revenue Trend</div>',
    unsafe_allow_html=True
)


if (
    detected["date"] is not None
    and
    detected["revenue"] is not None
):

    trend_df = filtered_df.copy()

    trend_df["Month"] = (
        trend_df[
            detected["date"]
        ]
        .dt.to_period("M")
        .astype(str)
    )


    monthly_revenue = (
        trend_df
        .groupby("Month")
        [detected["revenue"]]
        .sum()
        .reset_index()
    )


    fig_revenue = px.line(
        monthly_revenue,
        x="Month",
        y=detected["revenue"],
        markers=True,
        title="Monthly Revenue Trend"
    )


    fig_revenue.update_layout(
        xaxis_title="Month",
        yaxis_title="Revenue",
        hovermode="x unified"
    )


    st.plotly_chart(
        fig_revenue,
        use_container_width=True
    )

else:

    st.warning(
        "Date and revenue columns are required for the revenue trend."
    )


# ================================================================
# 29. REGION ANALYSIS
# ================================================================

st.markdown(
    '<div class="section-title">🌍 Regional Performance</div>',
    unsafe_allow_html=True
)


region_chart_col1, region_chart_col2 = st.columns(2)


if (
    detected["region"] is not None
    and
    detected["revenue"] is not None
):

    regional_revenue = (
        filtered_df
        .groupby(
            detected["region"]
        )[detected["revenue"]]
        .sum()
        .reset_index()
        .sort_values(
            detected["revenue"],
            ascending=False
        )
    )


    with region_chart_col1:

        fig_region = px.bar(
            regional_revenue,
            x=detected["region"],
            y=detected["revenue"],
            title="Revenue by Region",
            text_auto=".2s"
        )


        fig_region.update_layout(
            xaxis_title="Region",
            yaxis_title="Revenue"
        )


        st.plotly_chart(
            fig_region,
            use_container_width=True
        )


# ================================================================
# 30. USER ANALYSIS BY REGION
# ================================================================

if (
    detected["region"] is not None
    and
    detected["user"] is not None
):

    regional_users = (
        filtered_df
        .groupby(
            detected["region"]
        )[detected["user"]]
        .nunique()
        .reset_index()
    )


    regional_users.columns = [
        "Region",
        "Active Users"
    ]


    with region_chart_col2:

        fig_users = px.bar(
            regional_users,
            x="Region",
            y="Active Users",
            title="Active Users by Region",
            text_auto=True
        )


        fig_users.update_layout(
            xaxis_title="Region",
            yaxis_title="Active Users"
        )


        st.plotly_chart(
            fig_users,
            use_container_width=True
        )


# ================================================================
# 31. CATEGORY ANALYSIS
# ================================================================

st.markdown(
    '<div class="section-title">🏷️ Category Performance</div>',
    unsafe_allow_html=True
)


category_col1, category_col2 = st.columns(2)


if (
    detected["category"] is not None
    and
    detected["revenue"] is not None
):

    category_revenue = (
        filtered_df
        .groupby(
            detected["category"]
        )[detected["revenue"]]
        .sum()
        .reset_index()
        .sort_values(
            detected["revenue"],
            ascending=False
        )
    )


    with category_col1:

        fig_category = px.bar(
            category_revenue,
            x=detected["category"],
            y=detected["revenue"],
            title="Revenue by Category",
            text_auto=".2s"
        )


        fig_category.update_layout(
            xaxis_title="Category",
            yaxis_title="Revenue"
        )


       