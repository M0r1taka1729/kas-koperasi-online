import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# Inisialisasi koneksi
conn = st.connection("gsheets", type=GSheetsConnection)

# Saat membaca data
try:
    # Gunakan parameter ttl=0 agar data selalu paling baru (tidak tersimpan di cache)
    df = conn.read(skiprows=5, ttl=0)
    st.write("Data Berhasil Terhubung!")
    st.dataframe(df.tail())
except Exception as e:
    st.error(f"Koneksi Gagal: {e}")
