import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime
from fpdf import FPDF

st.set_page_config(page_title="Kas Koperasi Online", layout="wide")
st.title("🌐 Buku Kas Koperasi Online")

# Koneksi menggunakan Service Account yang ada di Secrets
conn = st.connection("gsheets", type=GSheetsConnection)

# --- FUNGSI PDF ---
def export_to_pdf(dataframe):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "LAPORAN BUKU KAS KOPERASI", ln=True, align='C')
    pdf.ln(5)
    widths = [25, 30, 35, 85, 30, 30, 35] 
    headers = ["Bulan", "Tanggal", "No Bukti", "Keterangan", "Debit", "Kredit", "Saldo"]
    pdf.set_fill_color(200, 220, 255)
    for i, h in enumerate(headers):
        pdf.cell(widths[i], 10, h, border=1, fill=True, align='C')
    pdf.ln()
    pdf.set_font("Arial", size=9)
    for index, row in dataframe.iterrows():
        fill = index % 2 == 0
        pdf.set_fill_color(245, 245, 245)
        for j in range(7):
            val = str(row.iloc[j]) if pd.notnull(row.iloc[j]) else ""
            pdf.cell(widths[j], 8, val[:45], border=1, fill=fill)
        pdf.ln()
    return pdf.output(dest='S').encode('latin-1')

# --- FORM INPUT ---
with st.form("form_online"):
    c1, c2 = st.columns(2)
    with c1:
        tgl = st.date_input("Tanggal", datetime.now())
        bukti = st.text_input("No. Bukti")
        ket = st.text_input("Keterangan")
    with c2:
        jenis = st.selectbox("Jenis", ["Debit", "Kredit"])
        nominal = st.number_input("Nominal", min_value=0)
    submit = st.form_submit_button("Simpan ke Cloud")

if submit:
    try:
        # 1. Ambil data dari Google Sheets
        # Kita baca mulai baris 5 agar judul terbaca sebagai header
        df_old = conn.read(skiprows=4, ttl=0)
        
        # 2. Pastikan data tidak kosong dan bersihkan baris yang isinya judul
        # Menghapus baris jika kolom 'Saldo' berisi teks 'Saldo'
        df_old = df_old[df_old.iloc[:, 6] != "Saldo"]
        
        # 3. Ambil Saldo Terakhir
        if not df_old.empty:
            # Ambil nilai dari baris terakhir, kolom ke-7 (indeks 6)
            last_val = df_old.iloc[-1, 6]
            # Paksa jadi angka, jika gagal (teks) jadikan 0
            try:
                last_saldo = float(last_val)
            except:
                last_saldo = 0
        else:
            last_saldo = 0
            
        # 4. Hitung Transaksi Baru
        deb = nominal if jenis == "Debit" else 0
        kre = nominal if jenis == "Kredit" else 0
        new_saldo = last_saldo + deb - kre
        
        # 5. Buat Baris Baru
        new_row = pd.DataFrame([{
            "Bulan": tgl.strftime("%B"),
            "Tanggal": tgl.strftime("%d/%m/%Y"),
            "No Bukti": bukti,
            "Keterangan": ket,
            "Debit": deb,
            "Kredit": kre,
            "Saldo": new_saldo
        }])
        
        # 6. Gabungkan dan Kirim kembali ke Cloud
        updated_df = pd.concat([df_old, new_row], ignore_index=True)
        conn.update(data=updated_df)
        
        st.success("✅ Data Berhasil Disimpan!")
        st.rerun()
        
    except Exception as e:
        st.error(f"Terjadi kesalahan saat menghitung saldo: {e}")
# --- DOWNLOAD & PREVIEW ---
st.divider()
try:
    data_cloud = conn.read(skiprows=5)
    if not data_cloud.empty:
        pdf_bytes = export_to_pdf(data_cloud)
        st.download_button("📄 Unduh PDF", pdf_bytes, "Laporan.pdf")
        st.dataframe(data_cloud.tail(5))
except:
    st.info("Menunggu data...")

