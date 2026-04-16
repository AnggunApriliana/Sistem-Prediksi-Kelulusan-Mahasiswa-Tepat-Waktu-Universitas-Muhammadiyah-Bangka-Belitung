import streamlit as st
import pandas as pd
import numpy as np
import joblib
import plotly.express as px
from datetime import datetime
import base64
import io

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="Executive Dashboard | UNMUH BABEL",
    page_icon="🎓",
    layout="wide"
)

# ==========================================
# 2. FUNCTION UTILITY (DOWNLOAD & CSS)
# ==========================================
def get_base64_file(file):
    try:
        with open(file, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None

def create_template():
    kolom_template = [
        "Nama", "Jenis Kelamin", "Program Studi", "Tahun Masuk", "Semester", 
        "IPK", "IPS 1", "IPS 2", "IPS 3", "IPS 4", "IPS 5", "IPS 6", "IPS 7", 
        "Jumlah SKS", "Jumlah Mata Kuliah yang Diulang", "Motivasi Belajar", 
        "Dukungan Keluarga", "Tingkat Stres", "Sosial-Ekonomi", 
        "Pekerjaan Paruh Waktu", "Keaktifan dalam Berorganisasi"
    ]
    df_template = pd.DataFrame(columns=kolom_template)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#1e3a8a', 'font_color': 'white', 'border': 1, 'align': 'center'
        })
        for col_num, value in enumerate(df_template.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 22)
    return buffer.getvalue()

banner_64 = get_base64_file("campus.jpg")
logo_64 = get_base64_file("logo.png")

banner_css = f"""
    background-image: linear-gradient(rgba(10,25,47,0.5), rgba(10,25,47,0.6)), url('data:image/jpg;base64,{banner_64}');
    background-size: cover; background-position: center center; min-height: 400px;
""" if banner_64 else "background-color: #0a192f;"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    .stApp {{ background-color: #f8fafc; font-family: 'Inter', sans-serif; }}
    
    .main-header {{
        {banner_css}
        padding: 60px 20px; border-radius: 0 0 40px 40px; color: white;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.1); margin-bottom: 40px;
    }}
    .header-content h1 {{ font-size: 2.8rem; font-weight: 800; margin-top: 15px; text-shadow: 2px 2px 8px rgba(0,0,0,0.7); }}
    
    .guide-card {{
        background: white; padding: 25px; border-radius: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); border-left: 6px solid #1e3a8a;
        height: 100%;
    }}
    .step-number {{
        background: #1e3a8a; color: white; width: 28px; height: 28px;
        border-radius: 50%; display: inline-flex; align-items: center;
        justify-content: center; font-weight: bold; margin-right: 10px; font-size: 0.9rem;
    }}
    
    .stat-card {{
        background: white; padding: 25px; border-radius: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); text-align: center; border-bottom: 5px solid #1e3a8a;
    }}
    
    .stButton>button {{
        width: 100%; border-radius: 12px; background-color: #1e3a8a; color: white;
        font-weight: 600; height: 3.5em; transition: 0.3s;
    }}
    .stButton>button:hover {{ background-color: #3b82f6; transform: translateY(-2px); }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. HEADER UI
# ==========================================
logo_html = f'<img src="data:image/png;base64,{logo_64}" style="height: 110px;">' if logo_64 else ""
st.markdown(f"""
<div class="main-header">
    {logo_html}
    <div class="header-content">
        <h1>SISTEM PREDIKSI KELULUSAN MAHASISWA</h1>
        <p style="font-size: 1.1rem; opacity: 0.9; letter-spacing: 3px; font-weight: 600;">UNIVERSITAS MUHAMMADIYAH BANGKA BELITUNG</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 4. PANDUAN PENGISIAN DATA
# ==========================================
st.markdown("### 📋 Panduan Pengisian Data")
g1, g2 = st.columns([1, 1.3])

with g1:
    st.markdown(f"""
    <div class="guide-card">
        <h4 style="color: #1e3a8a; margin-bottom: 20px;">🚀 Langkah Cepat</h4>
        <p><span class="step-number">1</span> Unduh <b>Template Excel</b> resmi.</p>
        <p><span class="step-number">2</span> Isi data sesuai dengan kolom yang tersedia.</p>
        <p><span class="step-number">3</span> Gunakan <b>huruf kecil</b> khusus untuk variabel perilaku/sosial.</p>
        <p><span class="step-number">4</span> Unggah kembali file pada kotak di bawah.</p>
        <div style="background-color: #fff4f4; padding: 12px; border-radius: 10px; border: 1px solid #fecaca; margin-top: 15px;">
            <p style="color: #b91c1c; font-size: 0.9rem; margin-bottom: 0;">
                <b>Note:</b> Sistem ini hanya bisa memprediksi mahasiswa aktif dari semester 1 sampai semester 8. 
                Untuk mahasiswa aktif semester 9 keatas tidak bisa diprediksi menggunakan sistem ini 🙏
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        label="📥 UNDUH TEMPLATE EXCEL KOSONG",
        data=create_template(),
        file_name="template_prediksi_unmuh.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with g2:
    st.markdown('<div class="guide-card">', unsafe_allow_html=True)
    st.markdown('<h4 style="color: #1e3a8a; margin-bottom: 15px;">🔍 Ketentuan Pilihan Jawaban</h4>', unsafe_allow_html=True)
    
    petunjuk_data = {
        "Kategori Variabel": ["Nama & Prodi", "IPK & IPS", "IPS (Jika belum ada)", "Motivasi & Dukungan", "Tingkat Stres", "Sosial-Ekonomi", "Pekerjaan", "Organisasi"],
        "Isi yang Benar": [
            "Bebas (Boleh Huruf Kapital)",
            "Gunakan titik (Contoh: 3.50)", 
            "Boleh kosong (Sistem akan menghitung rata-rata)",
            "sangat rendah, rendah, sedang, tinggi, sangat tinggi", 
            "rendah, sedang, tinggi", 
            "rendah, menengah, tinggi", 
            "bekerja, tidak bekerja", 
            "aktif, tidak aktif"
        ]
    }
    st.table(pd.DataFrame(petunjuk_data))
    st.caption("⚠️ Khusus variabel Motivasi s/d Organisasi, gunakan huruf kecil semua.")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# ==========================================
# 5. LOAD MODEL
# ==========================================
@st.cache_resource
def load_model():
    try:
        model = joblib.load("model_random_forest.pkl")
        fitur = joblib.load("fitur_model.pkl")
        return model, fitur
    except: return None, None

model, fitur_sistem = load_model()

# ==========================================
# 6. UPLOAD & PROCESSING
# ==========================================
st.markdown("### 📂 Unggah File yang Sudah Diisi")
file_up = st.file_uploader("Pilih file Excel", type=["xlsx"], label_visibility="collapsed")

if file_up:
    try:
        df_ori = pd.read_excel(file_up)
        df_ori.columns = df_ori.columns.str.strip()
        st.success(f"✔️ Berhasil mengimpor {len(df_ori)} data mahasiswa.")

        with st.expander("🔍 Klik untuk menampilkan data mentah"):
            st.dataframe(df_ori, use_container_width=True)
        
        if st.button("🚀 MULAI PROSES PREDIKSI"):
            with st.spinner("Sedang memprediksi..."):
                df_proc = df_ori.copy()
                
                # --- Preprocessing Kategori (Sesuai Training) ---
                mapping_5 = {'sangat rendah':1,'rendah':2,'sedang':3,'tinggi':4,'sangat tinggi':5}
                mapping_stres = {'rendah':1,'sedang':2,'tinggi':3}
                mapping_ekonomi = {'rendah':1,'menengah':2,'tinggi':3}
                mapping_kerja = {'tidak bekerja':0,'bekerja':1}
                mapping_org = {'tidak aktif':0,'aktif':1}

                cols_kategori = ['Motivasi Belajar', 'Dukungan Keluarga', 'Tingkat Stres', 
                                 'Sosial-Ekonomi', 'Pekerjaan Paruh Waktu', 'Keaktifan dalam Berorganisasi']

                for col in cols_kategori:
                    if col in df_proc.columns:
                        df_proc[col] = df_proc[col].astype(str).str.lower().str.strip()

                df_proc['Motivasi Belajar'] = df_proc['Motivasi Belajar'].map(mapping_5).fillna(3)
                df_proc['Dukungan Keluarga'] = df_proc['Dukungan Keluarga'].map(mapping_5).fillna(3)
                df_proc['Tingkat Stres'] = df_proc['Tingkat Stres'].map(mapping_stres).fillna(1)
                df_proc['Sosial-Ekonomi'] = df_proc['Sosial-Ekonomi'].map(mapping_ekonomi).fillna(2)
                df_proc['Pekerjaan Paruh Waktu'] = df_proc['Pekerjaan Paruh Waktu'].map(mapping_kerja).fillna(0)
                df_proc['Keaktifan dalam Berorganisasi'] = df_proc['Keaktifan dalam Berorganisasi'].map(mapping_org).fillna(0)

                # --- Penanganan Fitur (Sesuai Hasil Training: Tahun Masuk s/d IPS 11) ---
                if 'Tahun Lulus' not in df_proc.columns:
                    df_proc['Tahun Lulus'] = df_proc['Tahun Masuk'] + 4

                # Sesuaikan semua kolom IPS (1-11) yang dibutuhkan model
                all_ips = [f'IPS {i}' for i in range(1, 12)]
                for col in all_ips:
                    if col not in df_proc.columns:
                        df_proc[col] = np.nan
                
                # Mengisi nilai IPS yang kosong dengan rata-rata IPS yang ada (Imputasi adaptif)
                def fill_ips_adaptively(row):
                    existing = [row[c] for c in all_ips if c in row and not pd.isna(row[c]) and row[c] > 0]
                    avg = sum(existing)/len(existing) if existing else 0.0
                    for c in all_ips:
                        if pd.isna(row[c]) or row[c] <= 0:
                            row[c] = avg
                    return row

                df_proc = df_proc.apply(fill_ips_adaptively, axis=1)

                if 'IPK' in df_proc.columns: 
                    df_proc['IPK'] = df_proc['IPK'].clip(0, 4.0)

                # Reorder kolom sesuai urutan fitur_model.pkl hasil training
                X = df_proc.reindex(columns=fitur_sistem, fill_value=0)
                y_pred = model.predict(X)
                
                # Assign hasil ke dataframe original untuk ditampilkan
                df_ori['Estimasi Masa Studi'] = [round(float(val), 1) for val in y_pred]
                df_ori['Masa Studi'] = df_ori['Estimasi Masa Studi'].apply(lambda x: f"{x} Tahun")
                df_ori['Status'] = df_ori['Estimasi Masa Studi'].apply(lambda x: "TEPAT WAKTU" if x <= 4.0 else "TERLAMBAT")

                # --- Hasil Dashboard ---
                st.markdown("---")
                t1, t2 = st.columns(2)
                tepat = (df_ori['Status'] == "TEPAT WAKTU").sum()
                terlambat = (df_ori['Status'] == "TERLAMBAT").sum()

                with t1:
                    st.markdown(f'<div class="stat-card"><p style="color:#64748b">TEPAT WAKTU</p><h2 style="color:#1e3a8a">{tepat} Mahasiswa</h2></div>', unsafe_allow_html=True)
                with t2:
                    st.markdown(f'<div class="stat-card" style="border-bottom-color:#ef4444"><p style="color:#64748b">TERLAMBAT</p><h2 style="color:#ef4444">{terlambat} Mahasiswa</h2></div>', unsafe_allow_html=True)

                st.subheader("📋 Laporan Hasil Prediksi")
                df_final = df_ori[['Nama', 'Jenis Kelamin', 'Program Studi', 'Masa Studi', 'Status']]
                
                def color_status(val):
                    color = '#1e3a8a' if val == 'TEPAT WAKTU' else '#ef4444'
                    return f'color: {color}; font-weight: bold'

                styled_df = df_final.style.map(color_status, subset=['Status'])
                st.dataframe(styled_df, use_container_width=True)

                # --- Analisis Visual ---
                st.markdown("---")
                v1, v2, v3 = st.columns(3)
                
                with v1:
                    st.markdown("<p style='text-align: center; font-weight: 600;'>Status Kelulusan</p>", unsafe_allow_html=True)
                    fig_status = px.pie(df_ori, names='Status', hole=0.4, 
                                        color='Status', color_discrete_map={'TEPAT WAKTU': '#1e3a8a', 'TERLAMBAT': '#ef4444'})
                    fig_status.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=300)
                    st.plotly_chart(fig_status, use_container_width=True)
                
                with v2:
                    st.markdown("<p style='text-align: center; font-weight: 600;'>Jenis Kelamin</p>", unsafe_allow_html=True)
                    df_jk = df_ori['Jenis Kelamin'].value_counts().reset_index()
                    df_jk.columns = ['Jenis Kelamin', 'Jumlah']
                    fig_jk = px.bar(df_jk, x='Jenis Kelamin', y='Jumlah', color='Jenis Kelamin',
                                    color_discrete_sequence=['#1e3a8a', '#3b82f6'])
                    fig_jk.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=300, showlegend=False)
                    st.plotly_chart(fig_jk, use_container_width=True)

                with v3:
                    st.markdown("<p style='text-align: center; font-weight: 600;'>Program Studi</p>", unsafe_allow_html=True)
                    df_prodi = df_ori['Program Studi'].value_counts().reset_index()
                    df_prodi.columns = ['Program Studi', 'Jumlah']
                    fig_prodi = px.bar(df_prodi, y='Program Studi', x='Jumlah', orientation='h',
                                      color_discrete_sequence=['#10b981'])
                    fig_prodi.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=300)
                    st.plotly_chart(fig_prodi, use_container_width=True)

                # --- Unduh Hasil ---
                st.markdown("---")
                st.markdown("<h4 style='text-align: center;'>📥 Unduh Hasil Prediksi</h4>", unsafe_allow_html=True)
                
                excel_data = io.BytesIO()
                with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Hasil')
                
                _, db_col, _ = st.columns([1, 1, 1])
                with db_col:
                    st.download_button(
                        label="Download File Excel (.xlsx)",
                        data=excel_data.getvalue(),
                        file_name=f"prediksi_kelulusan_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("💡 Silakan unggah file Excel Anda untuk memulai prediksi.")

st.markdown(f"<div style='text-align:center; margin-top:50px; color:#94a3b8;'>© {datetime.now().year} UNIVERSITAS MUHAMMADIYAH BANGKA BELITUNG</div>", unsafe_allow_html=True)