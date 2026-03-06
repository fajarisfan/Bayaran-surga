import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Krakatau Baja Auto-Rekap", page_icon="🏗️")

st.title("🏗️ Rekapitulasi Karyawan & Mandays")
st.markdown("---")
st.write("Zizah, upload 12 file BA Bulanan di sini.")

days_input = st.number_input("Masukkan Jumlah Hari Kerja Efektif per Bulan", min_value=1, value=22)

uploaded_files = st.file_uploader("Pilih file Excel BA", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    rekap_data = []
    
    for file in uploaded_files:
        df = pd.read_excel(file)
        jumlah_orang = len(df.dropna(how='all'))
        total_mandays = jumlah_orang * days_input
        
        rekap_data.append({
            "Periode": file.name.split('.')[0],
            "Jumlah Karyawan": jumlah_orang,
            "Total Mandays": total_mandays
        })
    
    df_hasil = pd.DataFrame(rekap_data)
    st.subheader("📋 Hasil Tabel Rekapitulasi")
    st.table(df_hasil)
    
    # JURUS BIAR JADI TABEL EXCEL (.xlsx)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_hasil.to_excel(writer, index=False, sheet_name='Rekap_Zizah')
    
    st.download_button(
        label="📥 Download Hasil Rekap (Excel)",
        data=buffer.getvalue(),
        file_name='rekap_karyawan_mandays_zizah.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    st.success("Tugas 'Bayaran Surga' kelar! Zizah pasti happy. 😎")
