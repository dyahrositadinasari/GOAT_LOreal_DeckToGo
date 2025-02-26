import streamlit as st

st.title("GOAT-L'Oreal Monthly Report")
year = st.selectbox(
  'Please select the reporting month',
  ('2024', '2025', '2026')
)
month = st.selectbox(
  'Please select the reporting month',
  ('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')
)
st.write(
  "Year selected:", year,
  "\nMonth selected:", month
        )
