import streamlit as st

st.title("GOAT-L'Oreal Quarterly Report")
year = st.selectbox(
  'Please select the reporting year',
  ('2024', '2025')
)
quarter = st.selectbox(
  'Please select the reporting month',
  ('Q1', 'Q2', 'Q3', 'Q4')
)
quarter_map = {
    "Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4
}

quarter_num = quarter_map.get(quarter, "")  # Returns '' if quarter is not found

division = st.selectbox(
  "Please select the reporting L'Oreal Division",
  ('CPD', 'LDB', 'LLD', 'PPD')
)
category = st.selectbox(
  "Please select the reporting L'Oreal TDK Category",
  ('Hair Care', 'Female Skin', 'Make Up', 'Fragrance', 'Men Skin', 'Hair Color')
)
brands = st.multiselect(
    "Please Select 3 L'Oreal Brands to compare in the report",
    ["BLP Skin", "Garnier", "L'Oreal Paris", "GMN Shampoo Color", "Armani", "Kiehls", "Lancome", "Shu Uemura", "Urban Decay", "YSL", "Cerave", "La Roche Posay", "L'Oreal Professionel", "Matrix", "Biolage", "Kerastase", "Maybelline"]
,max_selections=3
)

from google.cloud import bigquery
from google.oauth2 import service_account
import json

# Load credentials from Streamlit Secrets
credentials_dict = st.secrets["gcp_service_account"]
credentials = service_account.Credentials.from_service_account_info(credentials_dict)

# Initialize BigQuery Client
client = bigquery.Client(credentials=credentials, project=credentials.project_id)

# Test connection
print(client.project)
