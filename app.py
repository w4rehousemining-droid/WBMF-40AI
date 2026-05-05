import streamlit as st

st.set_page_config(
    page_title="WBMF-40AI",
    page_icon="📦",
    layout="wide"
)

st.title("WBMF-40AI")
st.subheader("Warehouse Mining Business Management Framework - AI")

st.write("Halo! Jika Anda melihat ini, berarti setup Streamlit sudah berhasil.")

name = st.text_input("Siapa nama Anda?")

if name:
    st.success(f"Selamat datang di dunia Streamlit, {name}!")