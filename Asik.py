import streamlit as st
import pandas as pd

st.set_page_config(page_title="Krakatau Baja Auto-Rekap", page_icon="🏗️")

st.title("🏗️ Rekapitulasi Karyawan & Mandays")
st.markdown("---")
st.write("Zizah, upload 12 file BA Bulanan (Maret 2025 - Februari 2026) di sini.")

# Input jumlah hari kerja (default 22 hari/bulan)
days_input = st.number_input("Masukkan Jumlah Hari Kerja Efektif per Bulan", min_value=1, value=22)

# File uploader untuk banyak file sekaligus
uploaded_files = st.file_uploader("Pilih file Excel BA", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    rekap_data = []
    
    for file in uploaded_files:
        # Baca Excel (asumsi data nama ada di kolom pertama)
        df = pd.read_excel(file)
        
        # Bersihkan data kosong dan hitung jumlah orang
        jumlah_orang = len(df.dropna(how='all'))
        
        # Kalkulasi Mandays (Jumlah Orang x Hari Kerja)
        total_mandays = jumlah_orang * days_input
        
        rekap_data.append({
            "Periode": file.name.split('.')[0],
            "Jumlah Karyawan": jumlah_orang,
            "Total Mandays": total_mandays
        })
    
    # Tampilkan tabel hasil
    df_hasil = pd.DataFrame(rekap_data)
    st.subheader("📋 Hasil Tabel Rekapitulasi")
    st.table(df_hasil)
    
    # Tombol Download buat dipindahin ke Excel utama
    csv = df_hasil.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Hasil Rekap (CSV)",
        data=csv,
        file_name='rekap_karyawan_mandays_zizah.csv',
        mime='text/csv',
    )
    st.success("Tugas 'Bayaran Surga' kelar dalam hitungan detik! 😎")
