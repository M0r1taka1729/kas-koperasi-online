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
