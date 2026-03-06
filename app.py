import streamlit as st

st.title("Secrets-Test")

if "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
    st.success("Secrets sind vorhanden.")
    st.write("SUPABASE_URL gefunden.")
else:
    st.error("Secrets fehlen in der Cloud-App.")
    st.write("Gefundene Keys:", list(st.secrets.keys()))