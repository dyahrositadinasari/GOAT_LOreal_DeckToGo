import streamlit as st

#---PAGE SET-UP---
pages = {
    "Monthly Reporting L'Oreal": {
        "file": "pages/reporting_loreal_monthly.py",
        "icon": ":bar_chart:"
    },
    "Quarter Reporting L'Oreal": {
        "file": "pages/reporting_loreal_quarterly.py",
        "icon": ":chart_with_downwards_trend:"
    },
    "Annual Reporting L'Oreal": {
        "file": "pages/reporting_loreal_yearly.py",
        "icon": ":chart_with_upwards_trend:"
    }
}

#---NAVIGATION SET-UP---
st.sidebar.title("Navigation")
selection = st.sidebar.radio("Go to", list(pages.keys()))

#---RUN NAVIGATION---
selected_page = pages[selection]["file"]
exec(open(selected_page).read())
