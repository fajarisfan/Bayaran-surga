import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Krakatau Baja Master Rekap", page_icon="🏗️")

st.title("🏗️ Master Data Karyawan (Untuk VLOOKUP)")
st.markdown("---")
st.write("Upload semua file BA Bulanan, script ini akan menggabungkan **SEMUA NAMA** ke dalam satu file Excel.")

# Input jumlah hari kerja (opsional kalau mau tetep hitung mandays)
days_input = st.number_input("Masukkan Jumlah Hari Kerja Efektif per Bulan", min_value=1, value=22)

uploaded_files = st.file_uploader("Pilih file Excel BA", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    all_data_frames = [] 
    
    for file in uploaded_files:
        all_sheets = pd.read_excel(file, sheet_name=None)
        
        for sheet_name, df in all_sheets.items():
            # JURUS SAKTI: Cari baris mana yang mengandung kata "NAMA"
            mask = df.astype(str).apply(lambda x: x.str.contains('NAMA', case=False, na=False)).any(axis=1)
            
            if mask.any():
                header_idx = df[mask].index[0]
                
                # Ambil data lengkap (bukan cuma hitung jumlah)
                df_clean = pd.read_excel(file, sheet_name=sheet_name, skiprows=header_idx)
                
                # Identifikasi kolom Nama secara dinamis
                nama_cols = [c for c in df_clean.columns if 'NAMA' in str(c).upper()]
                
                if nama_cols:
                    col_key = nama_cols[0]
                    # Bersihkan baris yang namanya kosong (biar gak ada baris sampah)
                    df_clean = df_clean.dropna(subset=[col_key])
                    
                    # Tambahin identitas file & hitung mandays per baris (opsional)
                    df_clean['Sumber_File'] = file.name.split('.')[0]
                    df_clean['Mandays_Bulan_Ini'] = days_input
                    
                    all_data_frames.append(df_clean)

    if all_data_frames:
        # Gabungin SEMUA baris dari SEMUA file jadi satu
        df_master = pd.concat(all_data_frames, ignore_index=True)
        
        st.subheader(f"✅ Berhasil Mengumpulkan {len(df_master)} Baris Data")
        # Tampilkan tabel asli (Daftar Nama) di web
        st.dataframe(df_master.head(20)) 
        
        # Bikin file Excel Master
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_master.to_excel(writer, index=False, sheet_name='Master_Data_Zizah')
        
        st.download_button(
            label="📥 Download Master Data Lengkap (Excel)",
            data=buffer.getvalue(),
            file_name='MASTER_DATA_VLOOKUP_ZIZAH.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        st.success("Download file di atas, isinya sudah daftar nama lengkap untuk bahan VLOOKUP. 🥂")
    else:
        st.error("⚠️ Tidak ditemukan kolom 'NAMA'. Pastikan header di Excel tulisannya 'NAMA'.")

# --- ZONA TESTING (UNTUK PEMBUKTIAN) ---
st.markdown("---")
st.subheader("🛠️ Zona Testing (Versi Anti-Kosong)")
st.info("Download file di bawah, lalu upload kembali ke atas untuk tes.")

col1, col2 = st.columns(2)

# Logic Generator File Test yang BENER
def create_test_file(data, filename, start_row):
    buf = io.BytesIO()
    # Kita bikin DataFrame dulu
    df_test = pd.DataFrame(data)
    # Tulis ke Excel dengan 'startrow' untuk simulasi header berantakan
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_test.to_excel(writer, index=False, startrow=start_row, sheet_name='Sheet1')
    return buf.getvalue()

with col1:
    data_test1 = {
        'NO': [1, 2, 3],
        'NIK': ['A001', 'A002', 'A003'],
        'NAMA': ['Zizah Admin', 'Budi Teknik', 'Ani Logistik'],
        'JABATAN': ['Admin', 'Staf', 'Staf']
    }
    # Simulasi header 'NAMA' ada di baris ke-6 (startrow=5)
    file_a = create_test_file(data_test1, "Test_BA_Maret.xlsx", 5)
    st.download_button("📥 Download File Test A", file_a, "Test_BA_Maret.xlsx")

with col2:
    data_test2 = {
        'NO': [1, 2],
        'NIK': ['B001', 'B002'],
        'NAMA': ['Dedi Vendor', 'Eka Mandiri'],
        'JABATAN': ['Driver', 'Security']
    }
    # Simulasi header 'NAMA' ada di baris ke-4 (startrow=3)
    file_b = create_test_file(data_test2, "Test_BA_April.xlsx", 3)
    st.download_button("📥 Download File Test B", file_b, "Test_BA_April.xlsx")
