import streamlit as st
from google.oauth2 import service_account
from google.cloud import bigquery

# Create API client.
credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"]
)
client = bigquery.Client(credentials=credentials)


st.title("GOAT L'Oreal PPT Report Automation")
st.write("Please select the report type")
with st.sidebar:
    st.write("GOAT L'Oreal Monthly Report")


# Perform query.
# Uses st.cache_data to only rerun when the query changes or after 10 min.
@st.cache_data(ttl=600)
def run_query(query):
    query_job = client.query(query)
    rows_raw = query_job.result()
    # Convert to list of dicts. Required for st.cache_data to hash the return value.
    rows = [dict(row) for row in rows_raw]
    return rows

df = run_query("""
SELECT 
[Brand]
,SUM([Views]) as [Views]
,SUM([Engagement]) as [Engagement]
FROM loreal-id-prod.loreal_storage.advocacy_tdk
WHERE [TDK Category] = 'Female Skin'
AND MONTH([Date]) = '11'
AND YEAR([Date]) = '2024'
GROUP BY
[Brand]
""")

# Print results.
st.write(df)

