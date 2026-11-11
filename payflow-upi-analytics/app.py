import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# 1. Page Configuration
st.set_page_config(page_title="Payflow UPI Analytics", page_icon="💸", layout="wide")
st.title("💸 Payflow UPI Analytics Dashboard")

# 2. Data Loading (Replace with your actual data source)
@st.cache_data
def load_data():
    # Generating mock UPI data for demonstration
    dates = pd.date_range(start="2026-10-01", periods=30)
    df = pd.DataFrame({
        'Date': dates,
        'Transactions': np.random.randint(500, 2000, size=30),
        'Volume_INR': np.random.uniform(100000, 500000, size=30),
        'Success_Rate': np.random.uniform(92.0, 99.8, size=30)
    })
    return df

df = load_data()

# 3. Top Level KPIs
st.markdown("### Key Performance Indicators")
col1, col2, col3 = st.columns(3)

total_volume = df['Volume_INR'].sum()
total_txns = df['Transactions'].sum()
avg_success = df['Success_Rate'].mean()

col1.metric("Total Volume Processed", f"₹ {total_volume:,.0f}")
col2.metric("Total Transactions", f"{total_txns:,}")
col3.metric("Average Success Rate", f"{avg_success:.2f} %")

st.divider()

# 4. Visualizations
col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.subheader("Daily Transaction Volume")
    fig_vol = px.line(df, x='Date', y='Volume_INR', markers=True, 
                      labels={'Volume_INR': 'Volume (₹)'})
    st.plotly_chart(fig_vol, use_container_width=True)

with col_chart2:
    st.subheader("Network Success Rate")
    fig_sr = px.bar(df, x='Date', y='Success_Rate', 
                    color='Success_Rate', color_continuous_scale='Greens')
    fig_sr.update_yaxes(range=[85, 100]) # Zoom in on realistic UPI failure margins
    st.plotly_chart(fig_sr, use_container_width=True)

# 5. Raw Data Expander
with st.expander("View Raw UPI Data"):
    st.dataframe(df, use_container_width=True)
