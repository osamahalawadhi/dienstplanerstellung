import streamlit as st
from supabase import create_client


@st.cache_resource
def get_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


st.title("Supabase Verbindungstest")

sb = get_supabase()

st.subheader("Bestehende Planungsrunden")

try:
    result = sb.table("planning_rounds").select("*").order("id").execute()
    st.write(result.data)
except Exception as e:
    st.error("Lesen aus planning_rounds fehlgeschlagen.")
    st.exception(e)

st.markdown("---")
st.subheader("Neue Testrunde anlegen")

month = st.number_input("Monat", min_value=1, max_value=12, value=3, step=1)
year = st.number_input("Jahr", min_value=2025, max_value=2100, value=2026, step=1)
title = st.text_input("Titel", value="Dienstplan 03/2026")

if st.button("Testrunde speichern"):
    try:
        created = sb.table("planning_rounds").insert({
            "month": int(month),
            "year": int(year),
            "title": title
        }).execute()

        st.success("Testrunde gespeichert.")
        st.write(created.data)

    except Exception as e:
        st.error("Speichern fehlgeschlagen.")
        st.exception(e)

st.markdown("---")
if st.button("Planungsrunden neu laden"):
    st.rerun()