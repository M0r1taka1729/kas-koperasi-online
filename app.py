import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime
from fpdf import FPDF

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Kas Koperasi Online", layout="wide")
st.title("🌐 Input Kas Koperasi Online")

# --- KONEKSI KE GOOGLE SHEETS ---
# Tempelkan link Google Sheets Anda di sini
URL_SHEET = "https://docs.google.com/spreadsheets/d/1Os8nf0ourPsn1Rzt0yiMylNUX8Rkh1WwIjK736atMiU/edit?usp=sharing"

conn = st.connection("gsheets", type=GSheetsConnection)

# --- FUNGSI PDF ---
def export_to_pdf(dataframe):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "LAPORAN BUKU KAS KOPERASI ONLINE", ln=True, align='C')
    pdf.ln(10)
    
    widths = [25, 30, 35, 85, 30, 30, 35] 
    headers = ["Bulan", "Tanggal", "No Bukti", "Keterangan", "Debit", "Kredit", "Saldo"]
    
    pdf.set_fill_color(200, 220, 255)
    pdf.set_font("Arial", 'B', 10)
    for i, h in enumerate(headers):
        pdf.cell(widths[i], 10, h, border=1, fill=True, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", size=9)
    for index, row in dataframe.iterrows():
        fill = index % 2 == 0
        pdf.set_fill_color(245, 245, 245)
        for j in range(len(widths)):
            val = str(row.iloc[j]) if pd.notnull(row.iloc[j]) else ""
            # Jika kolom angka, format ribuan
            if j >= 4:
                try: val = f"{float(val):,.0f}"
                except: val = "0"
                pdf.cell(widths[j], 8, val, border=1, fill=fill, align='R')
            else:
                pdf.cell(widths[j], 8, val[:45], border=1, fill=fill)
        pdf.ln()
    return pdf.output(dest='S').encode('latin-1')

# --- FORM INPUT ---
with st.form("form_online"):
    col1, col2 = st.columns(2)
    with col1:
        tgl = st.date_input("Tanggal", datetime.now())
        bukti = st.text_input("No. Bukti")
        ket = st.text_input("Keterangan")
    with col2:
        jenis = st.selectbox("Jenis", ["Debit", "Kredit"])
        nominal = st.number_input("Nominal", min_value=0)
    
    submit = st.form_submit_button("Kirim ke Cloud")

if submit:
    # Ambil data lama dari Google Sheets
    existing_data = conn.read(spreadsheet=URL_SHEET, usecols=[0,1,2,3,4,5,6], skiprows=5)
    
    # Hitung Saldo Otomatis
    last_saldo = existing_data.iloc[-1, 6] if not existing_data.empty else 0
    debit = nominal if jenis == "Debit" else 0
    kredit = nominal if jenis == "Kredit" else 0
    new_saldo = float(last_saldo) + debit - kredit
    
    # Data baru
    new_row = pd.DataFrame([{
        "Bulan": tgl.strftime("%B"),
        "Tanggal": tgl.strftime("%d/%m/%Y"),
        "No Bukti": bukti,
        "Keterangan": ket,
        "Debit": debit,
        "Kredit": kredit,
        "Saldo": new_saldo
    }])
    
    # Gabung dan Update ke Google Sheets
    updated_df = pd.concat([existing_data, new_row], ignore_index=True)
    conn.update(spreadsheet=URL_SHEET, data=updated_df)
    st.success("Data Berhasil Masuk ke Google Sheets!")

# --- DOWNLOAD & PREVIEW ---
st.divider()
try:
    data_current = conn.read(spreadsheet=URL_SHEET, skiprows=5)
    st.write("Data Cloud Terbaru:")
    st.dataframe(data_current.tail(5))
    
    pdf_bytes = export_to_pdf(data_current)
    st.download_button("📥 Cetak PDF Laporan", pdf_bytes, "Laporan_Online.pdf")
except:
    st.info("Hubungkan ke Google Sheets untuk melihat data.")