import streamlit as st

st.title("GOAT-L'Oreal Monthly Report")
year = st.selectbox(
  'Please select the reporting year',
  ('2024', '2025')
)
month = st.selectbox(
  'Please select the reporting month',
  ('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')
)
division = st.selectbox(
  'Please select the reporting L'Oreal Division',
  ('CPD', 'LDB', 'LLD', 'PPD')
)
category = st.selectbox(
  'Please select the reporting L'Oreal TDK Category',
  ('Hair Care', 'Female Skin', 'Make Up', 'Fragrance', 'Men Skin', 'Hair Color')
)
brands = st.multiselect(
    'Please Select 3 Brands to compare in the report',
    ["BLP Skin", "Garnier", "L'Oreal Paris", "GMN Shampoo Color", "Armani", "Kiehls", "Lancome", "Shu Uemura", "Urban Decay", "YSL", "Cerave", "La Roche Posay", "L'Oreal Professionel", "Matrix", "Biolage", "Kerastase", "Maybelline"]
)
st.write("You selected:", brand)
