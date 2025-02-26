import streamlit as st

#---PAGE SET-UP---
m_page = st.Page(
    page ="pages/reporting_loreal_monthly.py",
    title ="Monthly Reporting L'Oreal",
    icon =":bar_chart:",
    default =True
)
q_page = st.Page(
    page ="pages/reporting_loreal_quarterly.py",
    title ="Quarter Reporting L'Oreal",
    icon =":chart_with_downwards_trend:",
)
a_page = st.Page(
    page ="pages/reporting_loreal_yearly.py",
    title ="Annual Reporting L'Oreal",
    icon =":chart_with_upwards_trend:",
)

#---NAVIGATION SET-UP---
pg = st.navigation(pages=[m_page, q_page, a_page])

#---RUN NAVIGATION---
pg.run()
