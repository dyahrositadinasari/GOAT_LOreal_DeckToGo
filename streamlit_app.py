import streamlit as st
import pyodbc as db

st.title("GOAT L'Oreal PPT Report Automation")
st.write("Please select the report type")
with st.sidebar:
    st.write("GOAT L'Oreal Monthly Report")
    
# Initialize connection.
# Uses st.cache_resource to only run once.
#@st.cache_resource
#def init_connection():
    #return 
conn = db.connect(
        "DRIVER={ODBC Driver 18 for SQL Server};SERVER="
        + st.secrets["server"]
        + ";DATABASE="
        + st.secrets["database"]
        + ";UID="
        + st.secrets["username"]
        + ";PWD="
        + st.secrets["password"]
    )

#conn = init_connection()

# Perform query.
# Uses st.cache_data to only rerun when the query changes or after 10 min.
@st.cache_data(ttl=600)
def run_query(query):
    with conn.cursor() as cur:
        cur.execute(query)
        return cur.fetchall()

rows = run_query("SELECT TOP 10 * FROM new_all_account_all_channel;")

# Print results.
for row in rows:
    st.write(f"{row[0]} has a :{row[1]}:")


