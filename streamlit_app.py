import streamlit as st
from google.oauth2 import service_account
from google.cloud import bigquery


st.title("GOAT L'Oreal PPT Report Automation")
st.write("Please select the report type")
with st.sidebar:
    st.write("GOAT L'Oreal Monthly Report")

#---PAGE SET-UP---
m_page = st.page(
    page ="pages/reporting_loreal_monthly.py",
    title ="Monthly Reporting L'Oreal"
    icon =":bar_chart:",
)
q_page = st.page(
    page ="pages/reporting_loreal_quarterly.py",
    title ="Quarter Reporting L'Oreal"
    icon =":chart_with_downwards_trend:",
)
a_page = st.page(
    page ="pages/reporting_loreal_yearly.py",
    title ="Annual Reporting L'Oreal"
    icon =":chart_with_upwards_trend:",
)

#---NAVIGATION SET-UP---
pg = st.navigation(pages=[m_page, q_page, y_page])

#---RUN NAVIGATION---
pg.run()
