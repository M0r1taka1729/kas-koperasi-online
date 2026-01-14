import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime
from fpdf import FPDF

# --- 1. KONFIGURASI ---
st.set_page_config(page_title="Kas Koperasi Online", layout="wide")
st.title("🌐 Buku Kas Koperasi Online (Google Sheets)")

# MASUKKAN LINK GOOGLE SHEETS ANDA DI SINI
URL_SHEET = "https://docs.google.com/spreadsheets/d/1Os8nf0ourPsn1Rzt0yiMylNUX8Rkh1WwIjK736atMiU/edit?usp=sharing"

# Koneksi ke Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

# --- 2. FUNGSI PDF ---
def export_to_pdf(dataframe):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 7, "KOPERASI KONSUMEN PENGAYOMAN LAPAS ARGAMAKMUR", ln=True, align='C')
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 7, "LAPORAN BUKU KAS ONLINE", ln=True, align='C')
    pdf.ln(5)
    
    # Lebar kolom & Header
    widths = [25, 30, 35, 85, 30, 30, 35] 
    headers = ["Bulan", "Tanggal", "No Bukti", "Keterangan", "Debit", "Kredit", "Saldo"]
    
    pdf.set_fill_color(230, 230, 230)
    pdf.set_font("Arial", 'B', 9)
    for i, h in enumerate(headers):
        pdf.cell(widths[i], 8, h, border=1, fill=True, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", size=9)
    for index, row in dataframe.iterrows():
        fill = index % 2 == 0
        pdf.set_fill_color(250, 250, 250)
        for j in range(7):
            val = str(row.iloc[j]) if pd.notnull(row.iloc[j]) else ""
            if j >= 4: # Kolom Angka
                try: val = f"{float(val):,.0f}"
                except: val = "0"
                pdf.cell(widths[j], 7, val, border=1, fill=fill, align='R')
            else:
                pdf.cell(widths[j], 7, val[:45], border=1, fill=fill)
        pdf.ln()
    return pdf.output(dest='S').encode('latin-1')

# --- 3. FORM INPUT ---
with st.form("form_online"):
    c1, c2 = st.columns(2)
    with c1:
        tgl = st.date_input("Tanggal", datetime.now())
        bukti = st.text_input("No. Bukti")
        ket = st.text_input("Keterangan")
    with c2:
        jenis = st.selectbox("Jenis", ["Debit", "Kredit"])
        nominal = st.number_input("Nominal", min_value=0)
    
    submit = st.form_submit_button("Simpan ke Google Sheets")

if submit:
    try:
        # Baca data lama (skiprows disesuaikan dengan struktur sheet Anda)
        df_old = conn.read(spreadsheet=URL_SHEET, skiprows=5)
        
        # Logika Saldo
        last_saldo = df_old.iloc[-1, 6] if not df_old.empty else 0
        deb = nominal if jenis == "Debit" else 0
        kre = nominal if jenis == "Kredit" else 0
        
        new_row = pd.DataFrame([{
            "Bulan": tgl.strftime("%B"),
            "Tanggal": tgl.strftime("%d/%m/%Y"),
            "No Bukti": bukti,
            "Keterangan": ket,
            "Debit": deb,
            "Kredit": kre,
            "Saldo": float(last_saldo) + deb - kre
        }])
        
        updated_df = pd.concat([df_old, new_row], ignore_index=True)
        conn.update(spreadsheet=URL_SHEET, data=updated_df)
        st.success("Data Berhasil Sinkron ke Cloud!")
        st.rerun()
    except Exception as e:
        st.error(f"Gagal Simpan: {e}")

# --- 4. TOMBOL CETAK & PREVIEW (SELALU MUNCUL) ---
st.divider()
st.subheader("📥 Cetak Laporan dari Cloud")

try:
    # Selalu ambil data terbaru untuk preview dan download
    data_cloud = conn.read(spreadsheet=URL_SHEET, skiprows=5)
    
    if not data_cloud.empty:
        col_pdf, col_link = st.columns(2)
        
        with col_pdf:
            pdf_bytes = export_to_pdf(data_cloud)
            st.download_button("📄 Unduh PDF Laporan", pdf_bytes, "Laporan_Kas_Online.pdf", "application/pdf")
            
        with col_link:
            st.link_button("📂 Buka Google Sheets", URL_SHEET)
            
        st.write("Preview 5 Data Terakhir:")
        st.dataframe(data_cloud.tail(5))
    else:
        st.info("Belum ada data di Google Sheets.")
except Exception as e:
    st.warning("Pastikan Link Google Sheets sudah benar dan aksesnya 'Anyone with the link can edit'.")
