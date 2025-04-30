import streamlit as st
import os
import pandas as pd
import numpy as np

import time
import json
from google.cloud import bigquery
from google.oauth2 import service_account
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_MARKER_STYLE
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.chart import XL_LABEL_POSITION
from pptx.enum.text import MSO_ANCHOR

st.title("GOAT-L'Oreal Quarterly Report")

st.title("zzzz Report")
st.badge("", icon="⚠️", color="red")
st.info("You can develop new report here")

year = st.selectbox(
  'Please select the reporting year',
  ('2024', '2025')
)
quarter = st.selectbox(
  'Please select the reporting quarter',
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
credentials = service_account.Credentials.from_service_account_info(st.secrets["gcp_service_account"])

# Initialize BigQuery Client
client = bigquery.Client(credentials=credentials, project=credentials.project_id)

# Test connection
st.write(print(client))
