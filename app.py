import streamlit as st
import pandas as pd
from datetime import datetime
import re
from io import BytesIO

st.set_page_config(page_title="OMT Tool", layout="wide")

st.title("📊 OMT Reporting Tool")
st.write("Upload your Excel file below")

# Upload
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"File loaded ✔ Rows: {len(df)}")
    st.dataframe(df.head())

# Date filter
col1, col2 = st.columns(2)

start_date = col1.date_input("Start Date")
end_date = col2.date_input("End Date")

# Run button
if st.button("Run Report"):
    if uploaded_file:
        st.success("Report running ✔ (basic version)")
    else:
        st.warning("Please upload a file first")
