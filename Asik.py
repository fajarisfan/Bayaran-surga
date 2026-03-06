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
    all_previews = [] 
    
    for file in uploaded_files:
        all_sheets = pd.read_excel(file, sheet_name=None)
        total_orang_file = 0
        file_samples = []
        
        for sheet_name, df in all_sheets.items():
            mask = df.astype(str).apply(lambda x: x.str.contains('NAMA', case=False, na=False)).any(axis=1)
            
            if mask.any():
                header_idx = df[mask].index[0]
                df_clean = pd.read_excel(file, sheet_name=sheet_name, skiprows=header_idx + 1)
                
                if len(df_clean.columns) >= 3:
                    target_col = df_clean.columns[2]
                    df_nama = df_clean[target_col].dropna()
                    
                    jumlah_orang = len(df_nama)
                    total_orang_file += jumlah_orang
                    
                    if not df_nama.empty:
                        samples = df_nama.head(3).astype(str).tolist()
                        file_samples.append(f"Sheet '{sheet_name}': {', '.join(samples)}...")
        
        total_mandays = total_orang_file * days_input
        rekap_data.append({
            "Periode": file.name.split('.')[0],
            "Jumlah Karyawan": total_orang_file,
            "Total Mandays": total_mandays
        })
        
        if file_samples:
            all_previews.append({"file": file.name, "samples": file_samples})
    
    if rekap_data:
        df_hasil = pd.DataFrame(rekap_data)
        st.subheader("📋 Hasil Tabel Rekapitulasi")
        st.table(df_hasil)
        
        with st.expander("🔍 Klik untuk Preview Isi Nama (Mastiin Gak Kosong)"):
            for item in all_previews:
                st.markdown(f"**📄 {item['file']}**")
                for s in item['samples']:
                    st.write(f"└─ {s}")
        
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
    else:
        st.warning("⚠️ File sudah diupload, tapi tidak ada kata 'NAMA' yang terdeteksi.")

# --- BAGIAN TAMBAHAN UNTUK DOWNLOAD FILE TEST ---
st.markdown("---")
st.subheader("🛠️ Zona Testing (Buat Zizah Coba-coba)")
st.write("Kalau gak ada file asli, klik tombol di bawah buat dapet file contoh.")

col1, col2 = st.columns(2)

with col1:
    # Generator File Normal
    buf1 = io.BytesIO()
    df_test1 = pd.DataFrame({'NO': [1,2,3], 'NIK': ['A1','A2','A3'], 'NAMA': ['Zizah','Budi','Ani']})
    # Kita sengaja kasih startrow=5 biar ada sampah di atasnya
    df_test1.to_excel(buf1, index=False, startrow=5, sheet_name='Data_Maret')
    
    st.download_button(
        label="📥 Download File Test 1 (Ada Sampah Header)",
        data=buf1.getvalue(),
        file_name="Test_Maret_Berantakan.xlsx"
    )

with col2:
    # Generator File Multi Sheet
    buf2 = io.BytesIO()
    with pd.ExcelWriter(buf2, engine='openpyxl') as writer:
        pd.DataFrame({'NO': [1,2], 'NAMA': ['Dedi','Eka']}).to_excel(writer, sheet_name='Grup_A', startrow=3, index=False)
        pd.DataFrame({'NO': [1,2,3], 'NAMA': ['Fani','Gita','Hani']}).to_excel(writer, sheet_name='Grup_B', startrow=2, index=False)
    
    st.download_button(
        label="📥 Download File Test 2 (Multi Sheet)",
        data=buf2.getvalue(),
        file_name="Test_Mei_MultiSheet.xlsx"
    )
