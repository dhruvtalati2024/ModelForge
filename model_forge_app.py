import streamlit as st
import os
from model import run_financial_analysis  # Import your core logic

st.set_page_config(page_title="Model Forge", layout="centered")

st.title("📊 Model Forge V3")
st.write("Generate clean, analyst-ready Excel models from any stock ticker.")

ticker = st.text_input("Enter Ticker Symbol (e.g., AAPL, TSLA, MSFT):")
years = st.slider("Select number of years of data:", min_value=3, max_value=10, value=5)

if st.button("Generate Excel Model"):
    if not ticker:
        st.warning("Please enter a ticker symbol.")
    else:
        with st.spinner("Fetching data and building Excel file..."):
            result = run_financial_analysis(ticker.upper(), years)
        if result:
            st.success("✅ Excel file generated!")
            with open(result, "rb") as f:
                st.download_button(
                    label="📥 Download Excel File",
                    data=f,
                    file_name=os.path.basename(result),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("Something went wrong. Please check the ticker and try again.")
