import streamlit as st

st.title("GOAT L'Oreal PPT Report Automation")
st.write("Please select the report type")

#---PAGE SET-UP---
pages = {
    "Monthly Reporting L'Oreal": {
        "file": "pages/reporting_loreal_monthly.py",
        "icon": "📊"
    },
    "Quarter Reporting L'Oreal": {
        "file": "pages/reporting_loreal_quarterly.py",
        "icon": "📈"
    },
    "Annual Reporting L'Oreal": {
        "file": "pages/reporting_loreal_yearly.py",
        "icon": "📉"
    }
}

#---NAVIGATION SET-UP---
#st.sidebar.title("Navigation")
#selection = st.sidebar.radio("Go to", list(pages.keys()))

#---RUN NAVIGATION---
#selected_page = pages[selection]["file"]
#exec(open(selected_page).read())
