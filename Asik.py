import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Krakatau Baja Auto-Rekap", page_icon="🏗️")

st.title("🏗️ Rekapitulasi Karyawan & Mandays")
st.markdown("---")
st.write("Zizah, upload file BA Bulanan (Maret 2025 - Februari 2026) di sini.")

# Input jumlah hari kerja
days_input = st.number_input("Masukkan Jumlah Hari Kerja Efektif per Bulan", min_value=1, value=22)

# File uploader
uploaded_files = st.file_uploader("Pilih file Excel BA", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    rekap_data = []
    
    for file in uploaded_files:
        # Baca semua sheet kalau filenya punya banyak sheet (kayak file absen tadi)
        all_sheets = pd.read_excel(file, sheet_name=None)
        
        total_orang_file = 0
        
        for sheet_name, df in all_sheets.items():
            # JURUS SAKTI: Cari baris mana yang mengandung kata "NAMA"
            # Ini biar header sampah di baris 1-9 otomatis kebuang
            mask = df.astype(str).apply(lambda x: x.str.contains('NAMA', case=False, na=False)).any(axis=1)
            
            if mask.any():
                header_idx = df[mask].index[0]
                # Ambil data setelah baris "NAMA" tersebut
                df_clean = pd.read_excel(file, sheet_name=sheet_name, skiprows=header_idx + 1)
                # Hitung baris yang kolom pertamanya gak kosong (asumsi itu kolom No/Nama)
                jumlah_orang = len(df_clean.dropna(subset=[df_clean.columns[2]])) # Ambil kolom ke-3 (Nama)
                total_orang_file += jumlah_orang
        
        # Kalkulasi Mandays
        total_mandays = total_orang_file * days_input
        
        rekap_data.append({
            "Periode": file.name.split('.')[0],
            "Jumlah Karyawan": total_orang_file,
            "Total Mandays": total_mandays
        })
    
    # Tampilkan tabel hasil
    df_hasil = pd.DataFrame(rekap_data)
    st.subheader("📋 Hasil Tabel Rekapitulasi")
    st.table(df_hasil)
    
    # Bikin file Excel (.xlsx) buat di-download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_hasil.to_excel(writer, index=False, sheet_name='Rekap_Mandays_Zizah')
    
    st.download_button(
        label="📥 Download Hasil Rekap (Excel)",
        data=buffer.getvalue(),
        file_name='rekap_karyawan_mandays_zizah.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    st.success("Logika 'Anti-Header Sampah' aktif. Zizah tinggal terima beres! 🥂")
