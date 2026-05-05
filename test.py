import streamlit as st

st.title("Aplikasi Streamlit Pertama Saya")
st.write("Halo! Jika Anda melihat ini, berarti setup Anda berhasil.")

name = st.text_input("Siapa nama Anda?")
if name:
    st.write(f"Selamat datang di dunia Streamlit, {name}!")