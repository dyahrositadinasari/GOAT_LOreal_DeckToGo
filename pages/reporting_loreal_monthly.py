import streamlit as st

st.title("GOAT-L'Oreal Monthly Report")
month = st.selectbox(
  'Please select the reporting month',
  ('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')
)
st.write('Month selected:', month)
