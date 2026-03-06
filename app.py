import streamlit as st
from supabase import create_client


@st.cache_resource
def get_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


st.title("Supabase Verbindungstest")

try:
    sb = get_supabase()
    result = sb.table("planning_rounds").select("*").execute()

    st.success("Verbindung zu Supabase erfolgreich.")
    st.write("Inhalt von planning_rounds:")
    st.write(result.data)

except Exception as e:
    st.error("Verbindung fehlgeschlagen.")
    st.exception(e)