import streamlit as st
#from streamlit_option_menu import option_menu

st.title("GOAT L'Oreal PPT Report Automation")
st.write("Please select the report type")

with st.sidebar:
    selected =option_menu(
        menu_title = "Select Report",
        options = ["Monthly Reporting LÓreal", "Quarter Reporting LÓreal", "Annual Reporting LÓreal"],
        icons = ["📊", "📈", "📉"],
        default_index = 0,
    )

