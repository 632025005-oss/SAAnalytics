# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import hashlib
import io
import re
import math
from openpyxl import Workbook
from streamlit_option_menu import option_menu
import pytz

# Import DB Storage & Helpers 
from db_storage import (
    load_users, save_users, login_siswa, simpan_survei_siswa, 
    get_data_gabungan, update_nilai_siswa, tambah_siswa_baru, 
    update_user_password, daftar_user_baru, get_semua_akun_siswa, 
    import_data_historis_lengkap, reset_password_siswa, hapus_survei_siswa
)

SUPER_ADMIN = {"admin", "regina"}

# ================================================================
# ★ FUNGSI PARSER
# ================================================================
def safe_int(val, default=0):
    if pd.isna(val): return default
    s = str(val).lower().replace('.', '')
    matches = re.findall(r'\d+', s)
    if len(matches) >= 2: return int((int(matches[0]) + int(matches[1])) / 2)
    elif len(matches) == 1: return int(matches[0])
    return default

def safe_float(val, default=0.0):
    if pd.isna(val): return default
    s = str(val).lower()
    matches = re.findall(r'\d+\.?\d*', s)
    if len(matches) >= 2: return (float(matches[0]) + float(matches[1])) / 2.0
    elif len(matches) == 1: return float(matches[0])
    return default

def safe_pendidikan(val):
    if pd.isna(val): return 3
    v = str(val).lower()
    if 'sd' in v: return 1
    if 'smp' in v: return 2
    if 'sma' in v or 'smk' in v: return 3
    if 'diploma' in v or 'sarjana' in v or 'd3' in v or 's1' in v or 'd1' in v or 'd2' in v: return 4
    if 's2' in v or 's3' in v or 'pasca' in v: return 5
    return 3

def safe_fasilitas(val):
    if pd.isna(val): return 0
    return len(str(val).split(','))

# ================================================================
# ★ KONSTANTA
# ================================================================
KONSTRUK_ITEMS = {
    "Self_Efficacy": [
        "Saya merasa puas ketika berhasil memahami materi yang sulit", 
        "Saya merasa nyaman dengan teman-teman sekelas", 
        "Saya yakin bisa memahami materi pelajaran yang paling sulit sekalipun", 
        "Saya yakin bisa mengerjakan tugas dengan baik jika saya berusaha keras", 
        "Saya yakin bisa menguasai keterampilan yang diajarkan di kelas", 
        "Saya yakin bisa mendapat nilai bagus jika saya belajar sungguh-sungguh"
    ],
    "Learning_Motivation": [
        "Saya senang belajar hal-hal baru di sekolah", 
        "Saya belajar karena ingin tahu lebih banyak", 
        "Saya belajar karena ingin mendapat nilai yang bagus", 
        "Saya belajar karena orang tua saya menyuruh", 
        "Saya merasa belajar di sekolah membuang waktu"
    ],
    "School_Support": [
        "Guru-guru menuntut saya untuk lulus dengan baik", 
        "Guru-guru di sekolah saya mengajar dengan baik", 
        "Guru memberikan perhatian jika saya kesulitan belajar", 
        "Suasana di sekolah saya mendukung untuk belajar"
    ],
    "Parental_Expectations": [
        "Orang tua mendorong saya untuk meraih cita-cita", 
        "Orang tua saya menuntut nilai yang tinggi"
    ],
    "School_Facilities": [
        "Fasilitas sekolah (perpustakaan, laboratorium) mendukung pembelajaran"
    ],
    "Academic_Anxiety": [
        "Saya merasa cemas menghadapi ujian akhir", 
        "Saya merasa tertekan dengan target nilai yang harus saya capai", 
        "Saya sulit memahami materi pelajaran"
    ],
    "Parental_Involvement": [
        "Orang tua menyediakan waktu untuk membantu saya belajar", 
        "Orang tua menghargai prestasi yang saya capai", 
        "Orang tua menyediakan fasilitas yang saya butuhkan untuk belajar", 
        "Orang tua memantau kegiatan belajar saya di rumah", 
        "Orang tua menanyakan tentang nilai atau tugas sekolah"
    ],
    "Learning_Laziness": [
        "Saya malas belajar"
    ]
}

REVERSE_ITEMS = {"Saya malas belajar", "Saya sulit memahami materi pelajaran", "Saya merasa belajar di sekolah membuang waktu"}

NON_LIKERT_ITEMS = {
    "Usia": {"type": "number", "min": 10, "max": 25, "default": 15, "step": 1, "label": "Usia (tahun)"},
    "Pendidikan Terakhir Ayah": {"type": "select", "options": [1, 2, 3, 4, 5], "default": 3, "labels": ["1 = SD", "2 = SMP", "3 = SMA/SMK", "4 = D3/S1", "5 = S2/S3"], "label": "Pendidikan Terakhir Ayah"},
    "Pendidikan Terakhir Ibu": {"type": "select", "options": [1, 2, 3, 4, 5], "default": 3, "labels": ["1 = SD", "2 = SMP", "3 = SMA/SMK", "4 = D3/S1", "5 = S2/S3"], "label": "Pendidikan Terakhir Ibu"},
    "Perkiraan Pendapatan Orang Tua per bulan": {"type": "number", "min": 0, "max": 50_000_000, "default": 2_000_000, "step": 500_000, "label": "Pendapatan Orang Tua/bulan (Rp)"},
    "Jarak rumah ke sekolah": {"type": "number", "min": 0.0, "max": 50.0, "default": 3.0, "step": 0.5, "label": "Jarak rumah–sekolah (km)"},
    "Fasilitas yang tersedia di rumah (centang semua yang sesuai):": {"type": "number", "min": 0, "max": 10, "default": 4, "step": 1, "label": "Jumlah fasilitas belajar di rumah"},
    "Rata-rata berapa jam per hari kamu menggunakan gadget/HP di luar keperluan belajar (bermain game, media sosial, menonton video)?": {"type": "number", "min": 0.0, "max": 24.0, "default": 2.5, "step": 0.5, "label": "Jam pakai HP non-belajar per hari"},
    "Rata-rata berapa jam kamu BELAJAR DI RUMAH setiap hari (di luar jam sekolah)?": {"type": "number", "min": 0.0, "max": 12.0, "default": 1.5, "step": 0.5, "label": "Jam belajar di rumah per hari"},
    "Berapa jam per MINGGU kamu mengikuti LES / BIMBINGAN BELAJAR di luar sekolah  ?": {"type": "number", "min": 0.0, "max": 40.0, "default": 0.0, "step": 0.5, "label": "Jam les/bimbel per minggu"},
    "Berapa banyak kegiatan EKSTRAKURIKULER yang kamu ikuti di sekolah selama kelas 9 semester 1?": {"type": "number", "min": 0, "max": 10, "default": 1, "step": 1, "label": "Jumlah ekstrakurikuler"},
}

BASELINE_ASPEK = {
    "Self_Efficacy": 4.634, 
    "Learning_Motivation": 2.517, 
    "School_Support": 4.600, 
    "Parental_Expectations": 3.450,
    "School_Facilities": 4.880,
    "Academic_Anxiety": 2.893, 
    "Parental_Involvement": 4.348, 
    "Learning_Laziness": 1.487
}

RATA_RATA_NILAI = 83.780

# --- DEFINISI MODEL DAG & GAMBARNYA ---
DAG_MODELS = {
    "Model 1: Pengaruh Langsung (Baseline)": {
        "deskripsi": "Model paling sederhana (Input-Output). Semua faktor dianggap mempengaruhi nilai secara langsung tanpa perantara.",
        "ate": {
            "Self_Efficacy": 1.500, "Learning_Motivation": -0.800, "School_Support": -0.500,
            "Parental_Expectations": -0.400, "School_Facilities": 0.300, "Academic_Anxiety": 0.600,
            "Parental_Involvement": 1.200, "Learning_Laziness": 0.500
        },
        "posisi": {
            "Nilai Raport": (0, 0),
            "Self-Efficacy": (-3, 2), "Motivation": (-1, 2), "School Support": (1, 2), "Parental Exp.": (3, 2),
            "School Facilities": (-3, -2), "Academic Anxiety": (-1, -2), "Parental Inv.": (1, -2), "Laziness": (3, -2)
        },
        "edges": [("Self-Efficacy", "Nilai Raport"), ("Motivation", "Nilai Raport"), ("School Support", "Nilai Raport"), ("Parental Exp.", "Nilai Raport"),
                   ("School Facilities", "Nilai Raport"), ("Academic Anxiety", "Nilai Raport"), ("Parental Inv.", "Nilai Raport"), ("Laziness", "Nilai Raport")]
    },
    "Model 2: Mediasi Penuh (Self-Determination)": {
        "deskripsi": "Lingkungan sosial (orang tua, sekolah) mempengaruhi prestasi melalui mekanisme psikologis (motivasi & efikasi diri).",
        "ate": {
            "Self_Efficacy": 2.100, "Learning_Motivation": 0.900, "School_Support": -0.100,
            "Parental_Expectations": -0.200, "School_Facilities": 0.400, "Academic_Anxiety": 0.800,
            "Parental_Involvement": 1.600, "Learning_Laziness": 0.600
        },
        "posisi": {
            "Nilai Raport": (0, 0),
            "Self-Efficacy": (-1, 1.5), "Motivation": (1, 1.5),
            # PERBAIKAN: Posisi variabel di atas digeser agar tidak bertumpuk
            "School Support": (-2.5, 3), "Parental Exp.": (0, 4), "Parental Inv.": (2.5, 3),
            "School Facilities": (-2, -1.5), "Academic Anxiety": (2, -1.5), "Laziness": (0, -2)
        },
        "edges": [("School Support", "Self-Efficacy"), ("Parental Inv.", "Self-Efficacy"), ("Parental Exp.", "Self-Efficacy"),
                   ("Parental Inv.", "Motivation"), ("School Support", "Motivation"),
                   ("Self-Efficacy", "Nilai Raport"), ("Motivation", "Nilai Raport"),
                   ("Laziness", "Nilai Raport"), ("Academic Anxiety", "Nilai Raport"), ("School Facilities", "Nilai Raport")]
    },
    "Model 3: Kecemasan sebagai Penekan (UTAMA)": {
        "deskripsi": "Model utama thesis (Lazarus & Bandura). Kecemasan berperan ganda: bisa memicu semangat, tapi juga bisa menekan jika berlebihan.",
        "ate": {
            "Self_Efficacy": +2.1995, "Parental_Involvement": +1.7559, "Learning_Laziness": +0.7320,
            "Academic_Anxiety": +0.7197, "School_Facilities": +0.6182, "Parental_Expectations": -0.6987,
            "School_Support": -1.0351, "Learning_Motivation": -1.9093
        },
        "posisi": {
            "Nilai Raport": (0, 0),
            "Cluster 0\n(Self-Efficacy)": (-2, 2), "Cluster 1\n(Motivation)": (-2, -1),
            "Cluster 2\n(School Support)": (0, 3), "Cluster 3\n(Parental Exp.)": (2, 3),
            "Cluster 4\n(School Facilities)": (-2, -3), "Cluster 5\n(Academic Anxiety)": (2, 0),
            "Cluster 6\n(Parental Inv.)": (0, -2), "Cluster 7\n(Laziness)": (2, -2)
        },
        "edges": [("Cluster 0\n(Self-Efficacy)", "Nilai Raport"), ("Cluster 1\n(Motivation)", "Nilai Raport"), ("Cluster 2\n(School Support)", "Nilai Raport"),
                   ("Cluster 3\n(Parental Exp.)", "Nilai Raport"), ("Cluster 4\n(School Facilities)", "Nilai Raport"), ("Cluster 5\n(Academic Anxiety)", "Nilai Raport"),
                   ("Cluster 6\n(Parental Inv.)", "Nilai Raport"), ("Cluster 7\n(Laziness)", "Nilai Raport")]
    },
    "Model 4: Model Lengkap (Ekologis & SRL)": {
        "deskripsi": "Model pemodelan komprehensif (Bronfenbrenner & Zimmerman). Mencakup semua jalur interaksi ekologis dan self-regulated learning.",
        "ate": {
            "Self_Efficacy": 2.500, "Learning_Motivation": 0.500, "School_Support": 0.200,
            "Parental_Expectations": -0.300, "School_Facilities": 0.700, "Academic_Anxiety": -0.500,
            "Parental_Involvement": 1.800, "Learning_Laziness": -0.900
        },
        "posisi": {
            "Nilai Raport": (0, 0),
            "Self-Efficacy": (-1, 1), "Motivation": (1, 1),
            "School Support": (-2, 3), "Parental Inv.": (2, 3), "Parental Exp.": (0, 4),
            "School Facilities": (-3, 1), "Academic Anxiety": (3, 1), "Laziness": (0, -2)
        },
        "edges": [("School Facilities", "Motivation"), ("School Facilities", "Self-Efficacy"), 
                   ("School Support", "Motivation"), ("School Support", "Self-Efficacy"),
                   ("Parental Inv.", "Motivation"), ("Parental Inv.", "Self-Efficacy"),
                   ("Parental Exp.", "Academic Anxiety"), ("Academic Anxiety", "Motivation"), ("Laziness", "Motivation"),
                   ("Self-Efficacy", "Nilai Raport"), ("Motivation", "Nilai Raport")]
    }
}

# ================================================================
# HELPER
# ================================================================
def hash_password(p): return hashlib.sha256(p.encode()).hexdigest()
def verify_password(p, hashed): return hash_password(p) == hashed

def compute_construct_scores(raw_answers: dict) -> dict:
    scores = {}
    for construct, items in KONSTRUK_ITEMS.items():
        vals = []
        for item in items:
            v = raw_answers.get(item, 3)
            if item in REVERSE_ITEMS: v = 6 - v
            vals.append(v)
        scores[construct] = sum(vals) / len(vals) if vals else 3.0
    return scores

def kategori_nilai(n):
    if n >= 88: return "Sangat Baik", "#10B981"
    if n >= 84: return "Baik", "#2563EB"
    if n >= 80: return "Cukup", "#F0B900"
    return "Perlu Perhatian", "#EF5B67"

def refresh_pengaruh():
    data_gabung = get_data_gabungan()
    st.session_state.database_siswa = data_gabung if data_gabung else []

# ================================================================
# PAGE CONFIG & CSS
# ================================================================
st.set_page_config(page_title="SAA Analytics", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
:root { --ink: #0f172a; --ink-2: #475569; --paper: #f8fafc; --card: #ffffff; --teal: #0d9488; --teal-deep: #134e4a; --line: #e2e8f0; --shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; color: var(--ink); }
.stApp { background-color: var(--paper); }
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1300px; }
#MainMenu, footer { visibility: hidden; }
h1, h2, h3, h4 { font-weight: 700; }
.top-header h1 { margin: 0; font-size: 2rem; color: var(--ink); font-weight: 700; }
.kpi-card { background: var(--card); border-radius: 8px; padding: 1.5rem; box-shadow: var(--shadow); border: 1px solid var(--line); display: flex; flex-direction: column; justify-content: center; }
.kpi-card h2 { margin: 0; font-size: 2rem; color: var(--ink); line-height: 1.2;}
.kpi-card p { margin: 0; font-size: 0.9rem; color: var(--ink-2); }
[data-testid="stSidebar"] { background-color: #0b132b !important; border-right: none; }
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
iframe[title="streamlit_option_menu.option_menu"] { background-color: transparent !important; }
.sb-logo { display: flex; align-items: center; gap: 15px; padding: 20px 20px 10px 20px; }
.sb-logo-text h2 { margin: 0; font-size: 1.5rem; font-weight: 700; color: white !important; }
.stButton > button { background: var(--teal); color: white; border-radius: 6px; font-weight: 600; border: none; padding: 0.5rem 1rem; transition: 0.2s; }
.stButton > button:hover { background: var(--teal-deep); color: white; }
div[data-testid="stForm"] { background: white; padding: 20px; border-radius: 10px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
</style>
""", unsafe_allow_html=True)

# ================================================================
# STATE MANAGEMENT
# ================================================================
for key, val in {
    "logged_in": False, "user_nama": None, "user_role": None,
    "database_siswa": [], "current_page": "home", "last_menu": "Beranda",
    "siswa_login_nisn": None, "edit_nisn": None, "preview_nisn": None,
    "page_survei": 1, "page_nilai": 1
}.items():
    if key not in st.session_state: st.session_state[key] = val

if st.session_state.logged_in: refresh_pengaruh()

# ================================================================
# FUNGSI RENDER KUESIONER
# ================================================================
def render_kuesioner_siswa(nisn, is_preview=False):
    if is_preview:
        st.info(f"👁️ **Mode Pratinjau:** Anda sedang melihat tampilan kuesioner siswa dengan NISN {nisn}.")
        if st.button("⬅️ Kembali ke Daftar Survei"):
            st.session_state.current_page = "list_survei" 
            st.session_state.preview_nisn = None
            st.rerun()
    else:
        st.success(f"✅ Login berhasil! NISN: {nisn}")
        
    st.markdown("### 📝 Kuesioner Siswa")
    df = pd.DataFrame(st.session_state.database_siswa)
    siswa_data = df[df['NISN'] == str(nisn)].iloc[0] if not df.empty and str(nisn) in df['NISN'].values else {}
    
    with st.form("form_kuesioner"):
        st.markdown("#### 🅰️ Bagian A — Identitas & Latar Belakang")
        
        c_id1, c_id2 = st.columns(2)
        default_nama = siswa_data.get("Nama", "") if isinstance(siswa_data, pd.Series) else ""
        default_kelas = siswa_data.get("Kelas", "") if isinstance(siswa_data, pd.Series) else ""
        
        nama_input = c_id1.text_input("Nama Lengkap Siswa", value=default_nama, placeholder="Tulis nama lengkap...")
        kelas_input = c_id2.text_input("Kelas (Contoh: IX-A)", value=default_kelas, placeholder="Contoh: 9E")
        
        non_likert_data = {}
        c1, c2 = st.columns(2)
        for i, (nama_kolom, cfg) in enumerate(NON_LIKERT_ITEMS.items()):
            col = c1 if i % 2 == 0 else c2
            val = siswa_data.get(nama_kolom, cfg["default"]) if isinstance(siswa_data, pd.Series) else cfg["default"]
            with col:
                if cfg["type"] == "number":
                    clean_val = safe_int(val, cfg["default"]) if isinstance(cfg["default"], int) else safe_float(val, cfg["default"])
                    # Perbaikan: hash() diganti md5 agar tidak reset saat server restart
                    non_likert_data[nama_kolom] = st.number_input(cfg["label"], min_value=cfg["min"], max_value=cfg["max"], value=clean_val, step=cfg.get("step", 1), key=f"nl_{hashlib.md5(nama_kolom.encode()).hexdigest()[:8]}")
                else:
                    idx_val = 2
                    if pd.notna(val):
                        try: idx_val = cfg["options"].index(int(val))
                        except: idx_val = cfg["options"].index(safe_pendidikan(val))
                    non_likert_data[nama_kolom] = st.selectbox(cfg["label"], options=cfg["options"], index=idx_val, format_func=lambda x, L=cfg["labels"], O=cfg["options"]: L[O.index(x)], key=f"nl_{hashlib.md5(nama_kolom.encode()).hexdigest()[:8]}")
        
        st.markdown("---")
        st.markdown("#### 🅱️ Bagian B — Pernyataan Psikologis")
        construct_answers = {}
        
        for idx, (construct, items) in enumerate(KONSTRUK_ITEMS.items(), 1):
            st.markdown(f"**{idx}. {construct.replace('_', ' ')}**")
            for j, item in enumerate(items, 1):
                val = siswa_data.get(construct, 3) if isinstance(siswa_data, pd.Series) else 3
                default_idx = int(round(safe_float(val, 3))) - 1 if pd.notna(val) else 2
                construct_answers[item] = st.radio(f"{item}", options=[1, 2, 3, 4, 5], horizontal=True, index=default_idx, key=f"q_{construct}_{j}")
            st.markdown("")
        
        if st.form_submit_button("📤 Kirim Kuesioner", type="primary", use_container_width=True):
            if not nama_input or not kelas_input:
                st.error("Mohon isi Nama Lengkap dan Kelas terlebih dahulu.")
            else:
                construct_scores = compute_construct_scores(construct_answers)
                payload = {**non_likert_data, **construct_scores, "Nama": nama_input, "Kelas": kelas_input}
                simpan_survei_siswa(nisn, payload)
                st.balloons()
                st.success("🎉 Kuesioner berhasil disimpan!")
                refresh_pengaruh()
                if is_preview:
                    st.session_state.current_page = "list_survei"
                    st.session_state.preview_nisn = None
                else:
                    st.session_state.siswa_login_nisn = None
                st.rerun()

    if not is_preview:
        if st.button("Batal / Logout Siswa"):
            st.session_state.siswa_login_nisn = None
            st.rerun()

# ================================================================
# HALAMAN LOGIN & REGISTER GURU
# ================================================================
def halaman_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br><h1 style='text-align:center; color:var(--teal); font-size: 2.8rem;'>🎓 SAA Analytics</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align:center; color:var(--ink-2); font-weight: 600; font-size: 1.1rem;'>Sistem Analitik Akademik Sekolah</p>", unsafe_allow_html=True)
        
        tab_login, tab_daftar, tab_siswa = st.tabs(["🔐 Login Guru/Admin", "📝 Daftar Akun Guru", "🎒 Portal Siswa"])
        
        with tab_login:
            with st.form("login_form"):
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                if st.form_submit_button("Masuk", use_container_width=True):
                    users = load_users()
                    if username in users:
                        if users[username].get("status") == "pending":
                            st.warning("Akun Anda masih menunggu persetujuan Admin.")
                        elif verify_password(password, users[username]["password"]):
                            st.session_state.logged_in = True
                            st.session_state.user_nama = users[username]["nama"]
                            st.session_state.user_role = users[username]["role"]
                            st.session_state.current_page = "home"
                            refresh_pengaruh()
                            st.rerun()
                        else: st.error("Password salah.")
                    else: st.error("Username tidak ditemukan.")

        with tab_daftar:
            st.info("Daftar akun Guru. Akun aktif setelah disetujui Admin.")
            with st.form("register_form"):
                reg_nama = st.text_input("Nama Lengkap")
                reg_user = st.text_input("Username (tanpa spasi)")
                reg_pass = st.text_input("Password Baru", type="password")
                if st.form_submit_button("Daftar Sekarang", type="primary", use_container_width=True):
                    if reg_nama and reg_user and reg_pass:
                        sukses, pesan = daftar_user_baru(reg_user, reg_nama, hash_password(reg_pass), "guru")
                        if sukses: st.success(pesan)
                        else: st.error(pesan)
                    else: st.error("Semua kolom wajib diisi.")

        with tab_siswa:
            if not st.session_state.siswa_login_nisn:
                st.info("Masukkan NISN. Jika belum punya akun, sistem akan mendaftarkan Anda otomatis (Password = NISN).")
                with st.form("form_login_siswa"):
                    nisn_input = st.text_input("NISN (Nomor Induk Siswa Nasional)")
                    pwd_input = st.text_input("Password (Ketik NISN Anda)", type="password")
                    if st.form_submit_button("Masuk / Daftar & Isi Kuesioner", type="primary"):
                        sukses, pesan = login_siswa(nisn_input, pwd_input)
                        if sukses:
                            df = pd.DataFrame(st.session_state.database_siswa)
                            siswa = df[df['NISN'] == str(nisn_input)]
                            if not siswa.empty and pd.notna(siswa.iloc[0].get('Self_Efficacy')):
                                st.warning("⚠️ Anda sudah pernah mengisi survei ini. Jika ada kesalahan, silakan hubungi Admin/Guru untuk menghapus dan mengulanginya.")
                            else:
                                st.session_state.siswa_login_nisn = nisn_input
                                st.rerun()
                        else: st.error(pesan)
            else:
                render_kuesioner_siswa(st.session_state.siswa_login_nisn, is_preview=False)

# ================================================================
# ROUTING UTAMA
# ================================================================
if not st.session_state.logged_in:
    if st.session_state.preview_nisn:
        render_kuesioner_siswa(st.session_state.preview_nisn, is_preview=True)
    else:
        halaman_login()
    st.stop()

# ================================================================
# SIDEBAR MENU
# ================================================================
with st.sidebar:
    st.markdown("""
    <div class="sb-logo">
        <div class="sb-logo-text">
            <h2>SAA<br><span style="font-size:1rem; font-weight:400;">Analytics</span></h2>
            <p style="font-size: 0.75rem; color: #94a3b8; margin-top: 5px;">Sistem Analitik Sekolah</p>
        </div>
    </div>
    <hr style="border-color: rgba(255,255,255,0.1); margin-top: 0;">
    """, unsafe_allow_html=True)

    menu_options = ["Beranda", "Siswa", "Jawaban Survei", "Nilai Raport", "Laporan Siswa", "Analisis Faktor", "Simulasi Nilai", "Akun", "Panduan"]
    menu_icons = ["house-door", "people", "clipboard-data", "bar-chart", "file-earmark-text", "diagram-3", "magic", "person-circle", "info-circle"]
    
    if st.session_state.user_role == "admin":
        menu_options.extend(["Kelola Pengguna", "Keluar"])
        menu_icons.extend(["gear", "box-arrow-right"])
    else:
        menu_options.append("Keluar")
        menu_icons.append("box-arrow-right")

    mapping_menu = {
        "Beranda": "home", "Siswa": "monitoring", "Jawaban Survei": "list_survei", 
        "Nilai Raport": "list_nilai", "Laporan Siswa": "rangkuman", "Analisis Faktor": "kausal", 
        "Simulasi Nilai": "prediksi", "Akun": "akun", "Panduan": "tentang", "Kelola Pengguna": "admin_users", 
        "Keluar": "logout"
    }

    current_page_mapped = st.session_state.current_page
    if current_page_mapped == "edit_x": current_page_mapped = "list_survei"
    if current_page_mapped == "edit_y": current_page_mapped = "list_nilai"

    default_idx = 0
    for i, opt in enumerate(menu_options):
        if mapping_menu.get(opt) == current_page_mapped:
            default_idx = i
            break

    menu_choice = option_menu(
        menu_title=None, options=menu_options, icons=menu_icons,
        default_index=default_idx, key="sidebar_menu",
        styles={
            "container": {"padding": "5px", "background-color": "rgba(0, 0, 0, 0.3)", "border-radius": "10px"},
            "icon": {"color": "#facc15", "font-size": "16px"},
            "nav-link": {"font-size": "15px", "text-align": "left", "color": "#e2e8f0 !important", "margin": "0px", "--hover-color": "rgba(255,255,255,0.1)"},
            "nav-link-selected": {"background-color": "rgba(255,255,255,0.15)", "color": "#facc15 !important", "border-left": "4px solid #facc15"},
        }
    )

    if menu_choice != st.session_state.last_menu:
        st.session_state.last_menu = menu_choice
        st.session_state.page_survei = 1
        st.session_state.page_nilai = 1
        
        if menu_choice == "Keluar":
            st.session_state.logged_in = False
            st.session_state.user_nama = None
            st.rerun()
        else:
            st.session_state.current_page = mapping_menu.get(menu_choice, "home")
        
    st.markdown("<br><br>", unsafe_allow_html=True) 
    st.markdown("---")
    st.markdown(f"""
    <div style="padding-left: 10px; padding-bottom: 20px;">
        <p style="color: #94a3b8; font-size: 0.8rem; margin-bottom: 0;">Masuk sebagai:</p>
        <p style="color: white; font-size: 0.95rem; font-weight: bold; margin-top: 2px;">👤 {st.session_state.user_nama}</p>
    </div>
    """, unsafe_allow_html=True)

# ================================================================
# HALAMAN: BERANDA
# ================================================================
if st.session_state.current_page == "home":
    st.markdown(f"""<div class="top-header" style="margin-bottom: 2rem;"><h1>Halo, {st.session_state.user_nama}</h1></div>""", unsafe_allow_html=True)

    refresh_pengaruh()
    df = pd.DataFrame(st.session_state.database_siswa)
    
    total_siswa = len(df)
    total_pertanyaan = sum(len(items) for items in KONSTRUK_ITEMS.values()) + len(NON_LIKERT_ITEMS)
    
    if not df.empty and 'Self_Efficacy' in df.columns:
        survei_lengkap = df['Self_Efficacy'].notna().sum()
        available_non_likert = [col for col in NON_LIKERT_ITEMS.keys() if col in df.columns]
        sel_kosong = df[df['Self_Efficacy'].notna()][available_non_likert].isna().sum().sum() if available_non_likert else 0
    else:
        survei_lengkap = 0; sel_kosong = 0
        
    if not df.empty and 'Nilai Akademik' in df.columns:
        nilai_terisi = (df['Nilai Akademik'] > 0).sum()
        y_kosong = total_siswa - nilai_terisi
    else:
        nilai_terisi = 0; y_kosong = 0

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="kpi-card"><h2>{total_siswa}</h2><p>Siswa</p></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><h2>{total_pertanyaan}</h2><p>Jumlah Pertanyaan</p></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><h2>{survei_lengkap} siswa</h2><p>Survei Lengkap</p></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card"><h2>{nilai_terisi} / {total_siswa}</h2><p>Nilai Terisi</p></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_chart, col_checklist = st.columns([2, 1])
    
    with col_chart:
        st.markdown("### Sebaran Nilai Siswa")
        if not df.empty and 'Nilai Akademik' in df.columns and (df['Nilai Akademik'] > 0).any():
            df_valid = df[df['Nilai Akademik'] > 0]
            fig = px.histogram(df_valid, x="Nilai Akademik", nbins=10, color_discrete_sequence=["#0d9488"])
            fig.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", xaxis_title="Rentang Nilai", yaxis_title="Jumlah Siswa", bargap=0.1)
            mean_val = df_valid['Nilai Akademik'].mean()
            fig.add_vline(x=mean_val, line_dash="dash", line_color="gray", annotation_text=f"Rata-rata {mean_val:.2f}")
            st.plotly_chart(fig, use_container_width=True)
            st.caption(f"Rata-rata nilai sekolah: **{mean_val:.2f}** | Standar Deviasi: **{df_valid['Nilai Akademik'].std():.2f}**")
        else: st.info("Belum ada data nilai untuk ditampilkan.")
            
    with col_checklist:
        st.markdown("""<div style="background: white; padding: 20px; border-radius: 8px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); height: 100%;"><h4 style="margin-top:0; font-size:1.1rem; color: #0f172a;">Data yang Belum Lengkap</h4>""", unsafe_allow_html=True)
        x_belum_isi = total_siswa - survei_lengkap
        st.markdown(f"""<ul style="list-style-type: none; padding-left: 0; font-size: 0.9rem; line-height: 2; color: #475569;"><li>📄 <b>{x_belum_isi} siswa</b> belum mengisi survei sama sekali</li><li>➖ <b>{sel_kosong} jawaban</b> masih kosong</li><li>⭕ <b>{y_kosong} siswa</b> belum punya nilai raport</li></ul><hr style="margin: 10px 0; border-color: #e2e8f0;"></div>""", unsafe_allow_html=True)

# ================================================================
# HALAMAN: MANAJEMEN SISWA
# ================================================================
elif st.session_state.current_page == "monitoring":
    st.title("Daftar & Manajemen Siswa")
    
    tab_tambah, tab_import, tab_daftar = st.tabs(["➕ Tambah Manual", "📥 Import Nilai (Excel)", "👥 Daftar Siswa"])
    
    with tab_tambah:
        with st.form("form_tambah_siswa"):
            c1, c2 = st.columns(2)
            with c1:
                new_nisn = st.text_input("NISN")
                new_nama = st.text_input("Nama Lengkap")
            with c2:
                new_kelas = st.text_input("Kelas")
                new_nilai = st.number_input("Nilai Raport", min_value=0.0, max_value=100.0, value=0.0, step=0.1)
            if st.form_submit_button("Simpan Siswa Baru", type="primary"):
                if new_nisn and new_nama and new_kelas:
                    tambah_siswa_baru(new_nisn, new_nama, new_kelas, new_nilai)
                    st.success(f"Data {new_nama} berhasil ditambahkan!")
                    refresh_pengaruh(); st.rerun()
                else: st.error("Semua kolom wajib diisi!")

    with tab_import:
        st.write("Unggah file Excel berisi Data Nilai (Dapodik). Sistem akan mencocokkan NISN dengan hasil survei siswa secara otomatis.")
        wb = Workbook(); ws = wb.active
        ws.append(["NISN", "Nama", "Kelas", "Nilai"])
        ws.append(["0051234567", "Ahmad Fauzi", "IX-A", 85.5])
        out = io.BytesIO(); wb.save(out)
        st.download_button("📄 Download Template Excel", data=out.getvalue(), file_name="Template_Nilai.xlsx")
        
        uploaded = st.file_uploader("Upload File Excel (.xlsx)", type=["xlsx"])
        if uploaded:
            try:
                df_upload = pd.read_excel(uploaded)
                st.dataframe(df_upload.head(), use_container_width=True)
                if st.button("Simpan Nilai ke Database", type="primary"):
                    count = import_data_historis_lengkap(df_upload)
                    st.success(f"✅ Berhasil! {count} data nilai berhasil dimasukkan dan dicocokkan dengan NISN siswa.")
                    refresh_pengaruh(); st.rerun()
            except Exception as e: st.error(f"Gagal memproses Excel: {e}")

    with tab_daftar:
        df_akun = get_semua_akun_siswa()
        if not df_akun.empty:
            csv_akun = df_akun.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download Daftar Akun Siswa (CSV)", data=csv_akun, file_name='Daftar_Akun_Siswa.csv', mime='text/csv')
            
        st.markdown("---")
        cari = st.text_input("🔍 Cari nama atau NISN siswa:")
        df_siswa = pd.DataFrame(st.session_state.database_siswa)
        if cari and not df_siswa.empty:
            mask = df_siswa['Nama'].astype(str).str.contains(cari, case=False, na=False) | df_siswa['NISN'].astype(str).str.contains(cari, case=False, na=False)
            df_siswa = df_siswa[mask]
            
        if df_siswa.empty: st.info("Belum ada data siswa yang cocok.")
        else:
            for idx, row in df_siswa.iterrows():
                col1, col2, col3, col4, col5, col6 = st.columns([1.5, 3, 2, 1, 2, 1.5])
                col1.write(f"**{row.get('NISN', '-')}**"); col2.write(f"**{row['Nama']}**"); col3.write(f"Kelas: {row.get('Kelas', '-')}"); col5.write(f"Nilai: {row.get('Nilai Akademik', '0')}")
                with col6:
                    b1, b2 = st.columns(2)
                    if b1.button("👁️ Portal", key=f"prev_{row['NISN']}"):
                        st.session_state.preview_nisn = row['NISN']; st.rerun()
                    if b2.button("✏️ Edit", key=f"edit_{row['NISN']}"):
                        st.session_state.edit_nisn = row['NISN']; st.session_state.current_page = "edit_x"; st.rerun()
                st.markdown("<hr style='margin: 5px 0; border-color: #e2e8f0;'>", unsafe_allow_html=True)

# ================================================================
# HALAMAN: DAFTAR SURVEI
# ================================================================
elif st.session_state.current_page == "list_survei":
    st.title("Data Survei Siswa")
    st.write("Klik tombol **Ubah** pada siswa untuk mengisi atau memperbarui jawaban survei.")
    
    refresh_pengaruh()
    df_siswa = pd.DataFrame(st.session_state.database_siswa)
    
    if df_siswa.empty: 
        st.warning("Belum ada data siswa.")
    else:
        cari = st.text_input("🔍 Cari nama atau NISN siswa:")
        if cari: 
            mask = df_siswa['Nama'].astype(str).str.contains(cari, case=False, na=False) | df_siswa['NISN'].astype(str).str.contains(cari, case=False, na=False)
            df_siswa = df_siswa[mask]
        
        st.markdown("---")
        
        items_per_page = 10
        total_pages = math.ceil(len(df_siswa) / items_per_page)
        
        if st.session_state.page_survei > total_pages: st.session_state.page_survei = total_pages
        if st.session_state.page_survei < 1: st.session_state.page_survei = 1
            
        start_idx = (st.session_state.page_survei - 1) * items_per_page
        end_idx = start_idx + items_per_page
        df_page = df_siswa.iloc[start_idx:end_idx]

        for i, (idx, row) in enumerate(df_page.iterrows()):
            main_constructs = ["Self_Efficacy", "Academic_Anxiety", "School_Support", "Parental_Involvement", "Learning_Motivation"]
            
            if all(pd.notna(row.get(c)) for c in main_constructs):
                status_teks = "37/37 (Lengkap) ✅"
            elif any(pd.notna(row.get(c)) for c in KONSTRUK_ITEMS.keys()):
                status_teks = "Sebagian ⚠️"
            else:
                status_teks = "0/37 (Belum) ❌"
            
            c1, c2, c3, c4, c5, c6 = st.columns([1.5, 4, 3, 1.2, 1.2, 1.2])
            c1.write(f"**{row['NISN']}**")
            c2.write(f"**{row['Nama']}**")
            c3.write(f"Terisi: {status_teks}")
            
            with c4:
                if st.button("Ubah", key=f"edit_survei_{row['NISN']}"):
                    st.session_state.edit_nisn = row['NISN']
                    st.session_state.current_page = "edit_x"
                    st.rerun()
            with c5:
                if st.button("🔑 Reset Pass", key=f"reset_pw_{row['NISN']}"):
                    reset_password_siswa(row['NISN'])
                    st.success(f"Password {row['Nama']} direset ke NISN.")
                    st.rerun()
            with c6:
                if st.button("🗑️ Hapus", key=f"del_survei_{row['NISN']}", type="primary"):
                    hapus_survei_siswa(row['NISN'])
                    st.success(f"Data survei {row['Nama']} dihapus. NISN dapat digunakan kembali.")
                    refresh_pengaruh(); st.rerun()
                    
            st.markdown("<hr style='margin: 5px 0; border-color: #e2e8f0;'>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- FITUR LOMPAT HALAMAN (PAGINATION) ---
        p1, p2, p3 = st.columns([1, 2, 1])
        with p1:
            if st.button("⬅️ Sebelumnya", disabled=(st.session_state.page_survei == 1), use_container_width=True):
                st.session_state.page_survei -= 1
                st.rerun()
        with p2:
            c_left, c_mid, c_right = st.columns([1, 1.5, 1])
            with c_mid:
                col_text1, col_select, col_text2 = st.columns([1, 1, 1])
                with col_text1:
                    st.markdown("<div style='text-align:right; padding-top:12px;'>Halaman</div>", unsafe_allow_html=True)
                with col_select:
                    selected_page_survei = st.selectbox(
                        "Lompat ke halaman", 
                        options=list(range(1, total_pages + 1)), 
                        index=st.session_state.page_survei - 1,
                        label_visibility="collapsed",
                        key="select_page_survei"
                    )
                    if selected_page_survei != st.session_state.page_survei:
                        st.session_state.page_survei = selected_page_survei
                        st.rerun()
                with col_text2:
                    st.markdown(f"<div style='text-align:left; padding-top:12px;'>dari <b>{total_pages}</b></div>", unsafe_allow_html=True)
        with p3:
            if st.button("Berikutnya ➡️", disabled=(st.session_state.page_survei == total_pages), use_container_width=True):
                st.session_state.page_survei += 1
                st.rerun()

# ================================================================
# HALAMAN: DAFTAR NILAI (PAGINATION)
# ================================================================
elif st.session_state.current_page == "list_nilai":
    st.title("Data Nilai Raport Siswa")
    st.write("Klik tombol **Ubah** pada siswa untuk memperbarui nilai raport.")
    
    df_siswa = pd.DataFrame(st.session_state.database_siswa)
    if df_siswa.empty: 
        st.warning("Belum ada data siswa.")
    else:
        cari = st.text_input("🔍 Cari nama atau NISN siswa:")
        if cari: 
            mask = df_siswa['Nama'].astype(str).str.contains(cari, case=False, na=False) | df_siswa['NISN'].astype(str).str.contains(cari, case=False, na=False)
            df_siswa = df_siswa[mask]
        
        st.markdown("---")
        
        items_per_page = 10
        total_pages = math.ceil(len(df_siswa) / items_per_page)
        
        if st.session_state.page_nilai > total_pages: st.session_state.page_nilai = total_pages
        if st.session_state.page_nilai < 1: st.session_state.page_nilai = 1
            
        start_idx = (st.session_state.page_nilai - 1) * items_per_page
        end_idx = start_idx + items_per_page
        df_page = df_siswa.iloc[start_idx:end_idx]

        for i, (idx, row) in enumerate(df_page.iterrows()):
            nilai_skrg = row.get('Nilai Akademik', 0)
            status_teks = "Lengkap" if nilai_skrg > 0 else "Kosong"
            
            c1, c2, c3, c4, c5 = st.columns([1.5, 4, 2, 2, 1])
            c1.write(f"**{row['NISN']}**"); c2.write(f"**{row['Nama']}**"); c3.write(f"Nilai: **{nilai_skrg:.2f}**"); c4.write(f"Status: {status_teks}")
            with c5:
                if st.button("Ubah", key=f"edit_nilai_{row['NISN']}"):
                    st.session_state.edit_nisn = row['NISN']
                    st.session_state.current_page = "edit_y"
                    st.rerun()
            st.markdown("<hr style='margin: 5px 0; border-color: #e2e8f0;'>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- FITUR LOMPAT HALAMAN (PAGINATION) ---
        p1, p2, p3 = st.columns([1, 2, 1])
        with p1:
            if st.button("⬅️ Sebelumnya", disabled=(st.session_state.page_nilai == 1), use_container_width=True, key="prev_nilai"):
                st.session_state.page_nilai -= 1
                st.rerun()
        with p2:
            c_left, c_mid, c_right = st.columns([1, 1.5, 1])
            with c_mid:
                col_text1, col_select, col_text2 = st.columns([1, 1, 1])
                with col_text1:
                    st.markdown("<div style='text-align:right; padding-top:12px;'>Halaman</div>", unsafe_allow_html=True)
                with col_select:
                    selected_page_nilai = st.selectbox(
                        "Lompat ke halaman", 
                        options=list(range(1, total_pages + 1)), 
                        index=st.session_state.page_nilai - 1,
                        label_visibility="collapsed",
                        key="select_page_nilai"
                    )
                    if selected_page_nilai != st.session_state.page_nilai:
                        st.session_state.page_nilai = selected_page_nilai
                        st.rerun()
                with col_text2:
                    st.markdown(f"<div style='text-align:left; padding-top:12px;'>dari <b>{total_pages}</b></div>", unsafe_allow_html=True)
        with p3:
            if st.button("Berikutnya ➡️", disabled=(st.session_state.page_nilai == total_pages), use_container_width=True, key="next_nilai"):
                st.session_state.page_nilai += 1
                st.rerun()

# ================================================================
# HALAMAN: EDIT SURVEI
# ================================================================
elif st.session_state.current_page == "edit_x":
    if st.session_state.get("edit_nisn") is None:
        st.warning("Pilih siswa terlebih dahulu di menu Jawaban Survei.")
        if st.button("⬅️ Kembali"):
            st.session_state.current_page = "list_survei"; st.session_state.last_menu = "Jawaban Survei"; st.rerun()
    else:
        df = pd.DataFrame(st.session_state.database_siswa)
        siswa = df[df['NISN'] == str(st.session_state.edit_nisn)].iloc[0]
        
        if st.button("⬅️ Kembali ke Daftar Survei"):
            st.session_state.current_page = "list_survei"; st.session_state.last_menu = "Jawaban Survei"; st.rerun()
            
        render_kuesioner_siswa(siswa['NISN'], is_preview=True)

# ================================================================
# HALAMAN: EDIT NILAI RAPORT
# ================================================================
elif st.session_state.current_page == "edit_y":
    if st.session_state.get("edit_nisn") is None:
        st.warning("Pilih siswa terlebih dahulu.")
        if st.button("⬅️ Kembali"):
            st.session_state.current_page = "list_nilai"; st.session_state.last_menu = "Nilai Raport"; st.rerun()
    else:
        df = pd.DataFrame(st.session_state.database_siswa)
        siswa = df[df['NISN'] == str(st.session_state.edit_nisn)].iloc[0]
        nisn = siswa['NISN']
        
        if st.button("⬅️ Kembali ke Daftar Nilai"):
            st.session_state.current_page = "list_nilai"; st.session_state.last_menu = "Nilai Raport"; st.rerun()
            
        nilai_skrg = float(siswa.get('Nilai Akademik', 0))
        
        st.markdown(f"""
        <div style="background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; border-left: 5px solid #0d9488;">
            <h2 style="margin: 0; color: #0f172a;">Data Y: {siswa['Nama']}</h2>
            <p style="margin: 5px 0 0 0; color: #475569;">Y saat ini: <b>{nilai_skrg:.3f}</b> (lengkap)</p>
        </div>
        """, unsafe_allow_html=True)
        
        mode_input = st.radio("Pilih cara pengisian nilai:", ["Satu angka rata-rata", "Nilai lengkap per mapel"], horizontal=True, label_visibility="collapsed")
        
        if mode_input == "Satu angka rata-rata":
            with st.form("form_nilai_rata"):
                st.markdown("<br>", unsafe_allow_html=True)
                nilai_tunggal = st.number_input("Masukkan Nilai Rata-rata Raport", min_value=0.0, max_value=100.0, value=nilai_skrg if nilai_skrg > 0 else 80.0, step=0.1)
                if st.form_submit_button("Simpan data Y", type="primary"):
                    sukses, pesan = update_nilai_siswa(nisn, {"Nilai": nilai_tunggal})
                    if sukses:
                        st.success(pesan); refresh_pengaruh(); st.session_state.current_page = "list_nilai"; st.session_state.last_menu = "Nilai Raport"; st.rerun()
                    else: st.error(pesan)
        else:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### Nilai lengkap per mapel")
            mapel_list = ["Pend. Agama", "PPKn", "B Indo", "Matematika", "B Ingg", "Seni Budaya/Prakarya", "PJOK", "B Jawa", "IPA", "IPS", "Informatika"]
            
            data_awal = []
            for m in mapel_list:
                s1 = safe_float(siswa.get(f"{m}_S1", nilai_skrg), nilai_skrg)
                s2 = safe_float(siswa.get(f"{m}_S2", nilai_skrg), nilai_skrg)
                data_awal.append({"Mata pelajaran": m, "Sem 1": s1, "Sem 2": s2, "Rata mapel": (s1+s2)/2})
            
            df_mapel = pd.DataFrame(data_awal)
            edited_df = st.data_editor(
                df_mapel, 
                column_config={
                    "Mata pelajaran": st.column_config.TextColumn(disabled=True),
                    "Sem 1": st.column_config.NumberColumn(min_value=0.0, max_value=100.0, step=0.01, format="%.2f"),
                    "Sem 2": st.column_config.NumberColumn(min_value=0.0, max_value=100.0, step=0.01, format="%.2f"),
                    "Rata mapel": st.column_config.NumberColumn(disabled=True, format="%.2f"),
                },
                hide_index=True, use_container_width=True
            )
            
            edited_df["Rata mapel"] = edited_df[["Sem 1", "Sem 2"]].mean(axis=1).round(2)
            total_rata = edited_df["Rata mapel"].mean()
            
            st.markdown(f"**Y = rata-rata 11 mapel: {total_rata:.3f}**")
            
            if st.button("Simpan data Y", type="primary"):
                update_data = {"Nilai": total_rata}
                for _, row in edited_df.iterrows():
                    update_data[f"{row['Mata pelajaran']}_S1"] = row["Sem 1"]
                    update_data[f"{row['Mata pelajaran']}_S2"] = row["Sem 2"]
                
                sukses, pesan = update_nilai_siswa(nisn, update_data)
                if sukses:
                    st.success(pesan); refresh_pengaruh(); st.session_state.current_page = "list_nilai"; st.session_state.last_menu = "Nilai Raport"; st.rerun()
                else: st.error(pesan)

# ================================================================
# HALAMAN: LAPORAN SISWA (PER SISWA)
# ================================================================
elif st.session_state.current_page == "rangkuman":
    st.title("Laporan Hasil Analisis Siswa")
    st.write("Pilih siswa untuk melihat profil, analisis faktor nilai, dan rekomendasi tindakan yang bisa dilakukan guru.")
    
    refresh_pengaruh()
    df = pd.DataFrame(st.session_state.database_siswa)
    
    if df.empty:
        st.warning("Belum ada data siswa.")
    else:
        df["Label"] = df["Nama"] + " (" + df["Kelas"] + ")"
        siswa_terpilih = st.selectbox("Pilih Siswa:", df["Label"].tolist())
        
        if siswa_terpilih:
            siswa = df[df["Label"] == siswa_terpilih].iloc[0]
            nisn = siswa["NISN"]
            nama = siswa["Nama"]
            kelas = siswa["Kelas"]
            nilai = safe_float(siswa.get("Nilai Akademik", 0))
            
            st.markdown("---")
            
            col_prof, col_val = st.columns([1, 1])
            with col_prof:
                st.markdown(f"""
                <div style="background: white; padding: 20px; border-radius: 8px; border: 1px solid #e2e8f0; border-left: 5px solid #0d9488;">
                    <h3 style="margin: 0; color: #0f172a;">{nama}</h3>
                    <p style="margin: 5px 0; color: #475569;">Kelas: {kelas} | NISN: {nisn}</p>
                </div>
                """, unsafe_allow_html=True)
            with col_val:
                kat, warna = kategori_nilai(nilai)
                st.markdown(f"""
                <div style="background: white; padding: 20px; border-radius: 8px; border: 1px solid #e2e8f0; border-left: 5px solid {warna};">
                    <h4 style="margin: 0; color: #475569;">Nilai Rata-rata</h4>
                    <h2 style="margin: 5px 0; color: {warna};">{nilai:.2f}</h2>
                    <p style="margin: 0; color: #475569;">Kategori: <b>{kat}</b></p>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            
            if nilai == 0:
                st.warning("Nilai siswa ini belum diinput oleh guru. Analisis mungkin tidak akurat.")
            
            st.markdown("### 📊 Mengapa Nilai Siswa Ini Seperti Ini?")
            st.write("Grafik di bawah menunjukkan hal-hal apa saja yang membuat nilai siswa ini **naik** atau **turun** dibandingkan rata-rata teman sekelasnya.")
            
            nama_awam = {
                "Self_Efficacy": "Percaya Diri", 
                "Learning_Motivation": "Semangat Belajar", 
                "School_Support": "Dukungan Sekolah", 
                "Parental_Expectations": "Harapan Orang Tua",
                "School_Facilities": "Fasilitas Sekolah",
                "Academic_Anxiety": "Rasa Cemas", 
                "Parental_Involvement": "Dukungan Orang Tua", 
                "Learning_Laziness": "Kemalasan Belajar"
            }
            
            # Menggunakan DAG 3 sebagai acuan default untuk Laporan
            current_ate = DAG_MODELS["Model 3: Kecemasan sebagai Penekan (UTAMA)"]["ate"]
            
            kontribusi = []
            for f, ate in current_ate.items():
                f_awam = nama_awam.get(f, f)
                val_siswa = safe_float(siswa.get(f, 3.0))
                baseline = BASELINE_ASPEK.get(f, 3.0)
                dampak = ate * (val_siswa - baseline)
                kontribusi.append({"Faktor": f_awam, "Pengaruh (Poin)": dampak})
                
            df_kontrib = pd.DataFrame(kontribusi).sort_values("Pengaruh (Poin)")
            
            fig = px.bar(df_kontrib, x="Pengaruh (Poin)", y="Faktor", orientation="h", color="Pengaruh (Poin)", color_continuous_scale="RdBu", text=[f"{v:+.2f} poin" for v in df_kontrib["Pengaruh (Poin)"]])
            fig.update_traces(textposition="outside")
            fig.update_layout(height=300, margin=dict(l=0, r=0, t=10, b=0), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            st.caption("💡 *Ke kanan (biru tua) berarti menambah nilai. Ke kiri (merah) berarti menurunkan nilai.*")

            st.markdown("### 💡 Rekomendasi Tindakan untuk Guru")
            
            faktor_penekan = df_kontrib.iloc[0] if len(df_kontrib) > 0 and df_kontrib.iloc[0]["Pengaruh (Poin)"] < 0 else None
            faktor_pendorong = df_kontrib.iloc[-1] if len(df_kontrib) > 0 and df_kontrib.iloc[-1]["Pengaruh (Poin)"] > 0 else None
            
            rekomendasi_teks = ""
            
            if faktor_penekan is not None:
                st.error(f"⚠️ **Hal yang paling menahan nilai siswa:** **{faktor_penekan['Faktor']}** (menurunkan nilai sekitar {abs(faktor_penekan['Pengaruh (Poin)']):.2f} poin).")
                rekomendasi_teks += f"Hal yang menahan nilai: {faktor_penekan['Faktor']}.\n"
                
                if "Percaya Diri" in faktor_penekan["Faktor"]:
                    st.markdown("- Berikan tugas yang bertahap, mulai dari yang mudah agar siswa merasa berhasil.\n- Puji usaha siswa, tidak hanya hasil akhirnya.\n- Ajak siswa melihat kembali hal-hal baik yang sudah ia pelajari minggu ini.")
                    rekomendasi_teks += "- Berikan tugas bertahap.\n- Puji usaha siswa.\n"
                elif "Rasa Cemas" in faktor_penekan["Faktor"]:
                    st.markdown("- Ajarkan cara menarik napas panjang saat siswa merasa panik sebelum ujian.\n- Berikan kuis kecil yang santai agar siswa terbiasa.\n- Beri tahu siswa bahwa rasa gugup itu wajar dan bisa membantunya lebih fokus.")
                    rekomendasi_teks += "- Ajarkan teknik relaksasi.\n- Berikan kuis santai.\n"
                elif "Dukungan Orang Tua" in faktor_penekan["Faktor"]:
                    st.markdown("- Hubungi orang tua untuk menyampaikan hal-hal positif tentang siswa (jangan hanya saat ada masalah).\n- Ajak orang tua berdiskusi ringan satu bulan sekali tentang perkembangan anaknya.")
                    rekomendasi_teks += "- Hubungi orang tua dengan kabar positif.\n- Ajak diskusi rutin.\n"
                elif "Semangat Belajar" in faktor_penekan["Faktor"]:
                    st.markdown("- Kaitkan materi pelajaran dengan hal yang disukai siswa.\n- Berikan apresiasi saat siswa menunjukkan minat pada suatu topik.\n- Ajak siswa menetapkan target belajar yang menyenangkan.")
                    rekomendasi_teks += "- Kaitkan materi dengan minat siswa.\n- Apresiasi minat siswa.\n"
                else:
                    st.markdown("- Ajak siswa mengobrol baik-baik untuk mencari tahu kendala apa yang sedang dihadapinya.")
                    rekomendasi_teks += "- Ajak diskusi mencari kendala.\n"
            
            if faktor_pendorong is not None:
                st.success(f"✅ **Hal yang sudah bagus dan menjadi kekuatan siswa:** **{faktor_pendorong['Faktor']}** (menambah nilai sekitar {faktor_pendorong['Pengaruh (Poin)']:.2f} poin).")
                rekomendasi_teks += f"Kekuatan siswa: {faktor_pendorong['Faktor']}.\n"
            else:
                st.info("Belum ada hal yang menonjol yang menjadi kekuatan siswa dibandingkan rata-rata teman-temannya.")
                
            st.markdown("---")
            laporan = f"""LAPORAN ANALISIS SISWA
Nama: {nama}
Kelas: {kelas}
NISN: {nisn}
Nilai Rata-rata: {nilai:.2f} ({kat})

Rekomendasi:
{rekomendasi_teks}
"""
            st.download_button(
                label="📥 Download Laporan Analisis Siswa (TXT)",
                data=laporan.encode('utf-8'),
                file_name=f"Laporan_{nama.replace(' ', '_')}.txt",
                mime="text/plain"
            )

# ================================================================
# HALAMAN: ANALISIS FAKTOR (PILIHAN MODEL DAG)
# ================================================================
elif st.session_state.current_page == "kausal":
    st.title("Analisis Faktor Nilai Siswa")
    st.write("Halaman ini menunjukkan alur hubungan faktor dan seberapa besar pengaruhnya terhadap nilai siswa.")
    
    model_terpilih = st.selectbox("Pilih Model Analisis Faktor (DAG):", list(DAG_MODELS.keys()))
    info_model = DAG_MODELS[model_terpilih]
    
    st.info(f"📌 **{model_terpilih}**: {info_model['deskripsi']}")
    
    mode_tampilan = st.radio("Pilih Tampilan Analisis:", ["Gambaran Umum Sekolah", "Penjelasan Satu Siswa"], horizontal=True)

    if mode_tampilan == "Gambaran Umum Sekolah":
        tab_alur, tab_tambah, tab_grafik = st.tabs(["🗺️ Alur Hubungan", "📝 Tambah Pertanyaan Baru", "📊 Pengaruh Tiap Faktor"])
        
        with tab_alur:
            st.markdown(f"### Alur Hubungan Antar Faktor ({model_terpilih})")
            st.write("Gambar ini menunjukkan hubungan sebab-akibat antar variabel. Garis panah menunjukkan arah pengaruh hingga akhirnya mempengaruhi Nilai Raport.")
            
            # --- GAMBAR DAG MENGGUNAKAN POSISI & EDGES DARI MODEL YANG DIPILIH ---
            posisi = info_model["posisi"]
            edges = info_model["edges"]
            
            edge_x, edge_y = [], []
            for e in edges:
                x0, y0 = posisi[e[0]]; x1, y1 = posisi[e[1]]
                edge_x.extend([x0, x1, None]); edge_y.extend([y0, y1, None])

            edge_trace = go.Scatter(x=edge_x, y=edge_y, line=dict(width=1.5, color='#888'), hoverinfo='none', mode='lines')
            node_x, node_y, node_text = [], [], []
            for node, (x, y) in posisi.items():
                node_x.append(x); node_y.append(y); node_text.append(node)

            node_trace = go.Scatter(x=node_x, y=node_y, mode='markers+text', text=node_text, textposition="top center", hoverinfo='text', marker=dict(showscale=False, color='#1f77b4', size=30, line_width=2))
            fig_dag = go.Figure(data=[edge_trace, node_trace], layout=go.Layout(showlegend=False, hovermode='closest', margin=dict(b=20,l=5,r=5,t=20), xaxis=dict(showgrid=False, zeroline=False, showticklabels=False), yaxis=dict(showgrid=False, zeroline=False, showticklabels=False), height=450, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)"))
            st.plotly_chart(fig_dag, use_container_width=True)

        with tab_tambah:
            st.markdown("### Tambah Pertanyaan / Faktor Baru")
            st.write("Gunakan menu ini jika ada pertanyaan baru yang ingin ditambahkan ke survei siswa.")
            with st.form("form_tambah_pertanyaan"):
                variabel_baru = st.text_input("Nama Faktor (Contoh: Jam Tidur)")
                pertanyaan_baru = st.text_area("Teks Pertanyaan untuk Siswa")
                tipe_jawaban = st.selectbox("Tipe Jawaban", ["Skala 1-5", "Angka Murni"])
                if st.form_submit_button("Simpan Pertanyaan Baru", type="primary"):
                    st.success(f"Faktor '{variabel_baru}' berhasil ditambahkan ke sistem!")
        
        with tab_grafik:
            st.markdown(f"### Seberapa Besar Pengaruh Faktor Ini ke Nilai? (Berdasarkan {model_terpilih})")
            st.write("Grafik di bawah menunjukkan, jika faktor ini ditingkatkan, apakah nilai siswa akan naik atau turun secara keseluruhan berdasarkan analisis kausal model yang dipilih.")
            
            current_ate = info_model["ate"]
            df_ate = pd.DataFrame({"Faktor": list(current_ate.keys()), "Pengaruh": list(current_ate.values())}).sort_values("Pengaruh")
            nama_awam = {
                "Self_Efficacy": "Percaya Diri", 
                "Learning_Motivation": "Semangat Belajar", 
                "School_Support": "Dukungan Sekolah", 
                "Parental_Expectations": "Harapan Orang Tua",
                "School_Facilities": "Fasilitas Sekolah",
                "Academic_Anxiety": "Rasa Cemas", 
                "Parental_Involvement": "Dukungan Orang Tua", 
                "Learning_Laziness": "Kemalasan Belajar"
            }
            df_ate["Faktor"] = df_ate["Faktor"].map(nama_awam)
            
            st.info(f"💡 **Kesimpulan:** Faktor yang paling kuat menaikkan nilai secara umum adalah **{df_ate.iloc[-1]['Faktor']}**.")
            
            fig = px.bar(df_ate, x="Pengaruh", y="Faktor", orientation="h", color="Pengaruh", color_continuous_scale="RdBu", text=[f"{v:+.2f}" for v in df_ate["Pengaruh"]])
            fig.update_traces(textposition="outside")
            fig.update_layout(height=400, margin=dict(l=0, r=0, t=10, b=0), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

    else: # Mode "Penjelasan Satu Siswa"
        refresh_pengaruh()
        df = pd.DataFrame(st.session_state.database_siswa)
        
        if df.empty:
            st.warning("Belum ada data siswa.")
        else:
            df["Label"] = df["Nama"] + " (" + df["Kelas"] + ")"
            siswa_terpilih = st.selectbox("Pilih Siswa untuk Dianalisis:", df["Label"].tolist())

            if siswa_terpilih:
                siswa = df[df["Label"] == siswa_terpilih].iloc[0]
                nama = siswa["Nama"]
                nisn = siswa["NISN"]
                kelas = siswa["Kelas"]
                nilai = safe_float(siswa.get("Nilai Akademik", 0))

                st.markdown("---")
                st.markdown(f"### Analisis Faktor: {nama} ({kelas})")
                st.write("Berikut adalah kesimpulan singkat mengenai faktor yang mempengaruhi nilai siswa ini berdasarkan model yang dipilih.")

                nama_awam = {
                    "Self_Efficacy": "Percaya Diri", 
                    "Learning_Motivation": "Semangat Belajar", 
                    "School_Support": "Dukungan Sekolah", 
                    "Parental_Expectations": "Harapan Orang Tua",
                    "School_Facilities": "Fasilitas Sekolah",
                    "Academic_Anxiety": "Rasa Cemas", 
                    "Parental_Involvement": "Dukungan Orang Tua", 
                    "Learning_Laziness": "Kemalasan Belajar"
                }

                current_ate = info_model["ate"]
                kontribusi = []
                for f, ate in current_ate.items():
                    f_awam = nama_awam.get(f, f)
                    val_siswa = safe_float(siswa.get(f, 3.0))
                    baseline = BASELINE_ASPEK.get(f, 3.0)
                    dampak = ate * (val_siswa - baseline)
                    kontribusi.append({"Faktor": f_awam, "Pengaruh (Poin)": dampak})

                df_kontrib = pd.DataFrame(kontribusi).sort_values("Pengaruh (Poin)")

                faktor_penekan = df_kontrib.iloc[0] if len(df_kontrib) > 0 and df_kontrib.iloc[0]["Pengaruh (Poin)"] < 0 else None
                faktor_pendorong = df_kontrib.iloc[-1] if len(df_kontrib) > 0 and df_kontrib.iloc[-1]["Pengaruh (Poin)"] > 0 else None

                st.markdown("<br>", unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    if faktor_pendorong is not None:
                        st.success(f"✅ **Faktor yang Membantu (Kekuatan):**\n\n**{faktor_pendorong['Faktor']}**\n\nFaktor ini membuat nilai siswa lebih baik {faktor_pendorong['Pengaruh (Poin)']:.2f} poin dibandingkan teman-temannya.")
                    else:
                        st.info("Belum ada faktor kekuatan yang menonjol.")
                with c2:
                    if faktor_penekan is not None:
                        st.error(f"⚠️ **Faktor yang Menghambat:**\n\n**{faktor_penekan['Faktor']}**\n\nFaktor ini membuat nilai siswa lebih rendah {abs(faktor_penekan['Pengaruh (Poin)']):.2f} poin dibandingkan teman-temannya.")
                    else:
                        st.info("Tidak ada faktor penghambat yang signifikan.")

                st.markdown("<br>", unsafe_allow_html=True)
                with st.expander("Lihat Detail Grafik (Opsional)"):
                    st.write("Grafik di bawah menunjukkan perbandingan faktor siswa ini dengan rata-rata teman sekelasnya.")
                    fig = px.bar(df_kontrib, x="Pengaruh (Poin)", y="Faktor", orientation="h", color="Pengaruh (Poin)", color_continuous_scale="RdBu", text=[f"{v:+.2f} poin" for v in df_kontrib["Pengaruh (Poin)"]])
                    fig.update_traces(textposition="outside")
                    fig.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", coloraxis_showscale=False)
                    st.plotly_chart(fig, use_container_width=True)

                laporan = f"LAPORAN ANALISIS FAKTOR SISWA\nNama: {nama}\nKelas: {kelas}\nNISN: {nisn}\nNilai: {nilai:.2f}\nModel: {model_terpilih}\n\nPengaruh Faktor:\n"
                for _, row in df_kontrib.iterrows():
                    laporan += f"- {row['Faktor']}: {row['Pengaruh (Poin)']:+.2f} poin\n"

                st.download_button(
                    label="📥 Download Laporan Analisis Faktor (TXT)",
                    data=laporan.encode('utf-8'),
                    file_name=f"Analisis_Faktor_{nama.replace(' ', '_')}.txt",
                    mime="text/plain"
                )

# ================================================================
# HALAMAN LAINNYA (PREDIKSI, AKUN, TENTANG, ADMIN)
# ================================================================
elif st.session_state.current_page == "prediksi":
    st.title("Simulasi Nilai Siswa")
    st.write("Geser nilai di bawah ini untuk melihat perkiraan nilai siswa jika faktor-faktornya diubah (Berdasarkan DAG 3).")
    
    c1, c2 = st.columns(2)
    with c1:
        p_se  = st.slider("Percaya Diri", 1.0, 5.0, BASELINE_ASPEK["Self_Efficacy"], 0.1)
        p_mot = st.slider("Semangat Belajar", 1.0, 5.0, BASELINE_ASPEK["Learning_Motivation"], 0.1)
        p_cem = st.slider("Rasa Cemas", 1.0, 5.0, BASELINE_ASPEK["Academic_Anxiety"], 0.1)
        p_lazy = st.slider("Kemalasan Belajar", 1.0, 5.0, BASELINE_ASPEK["Learning_Laziness"], 0.1)
    with c2:
        p_ortu = st.slider("Dukungan Orang Tua", 1.0, 5.0, BASELINE_ASPEK["Parental_Involvement"], 0.1)
        p_exp = st.slider("Harapan Orang Tua", 1.0, 5.0, BASELINE_ASPEK["Parental_Expectations"], 0.1)
        p_sekolah = st.slider("Dukungan Sekolah", 1.0, 5.0, BASELINE_ASPEK["School_Support"], 0.1)
        p_fas = st.slider("Fasilitas Sekolah", 1.0, 5.0, BASELINE_ASPEK["School_Facilities"], 0.1)
        
    if st.button("Hitung Perkiraan Nilai", type="primary"):
        profil = {
            "Self_Efficacy": p_se, "Learning_Motivation": p_mot, "Academic_Anxiety": p_cem, "Learning_Laziness": p_lazy,
            "Parental_Involvement": p_ortu, "Parental_Expectations": p_exp, "School_Support": p_sekolah, "School_Facilities": p_fas
        }
        current_ate = DAG_MODELS["Model 3: Kecemasan sebagai Penekan (UTAMA)"]["ate"]
        score = RATA_RATA_NILAI + sum([current_ate[k] * (v - BASELINE_ASPEK[k]) for k, v in profil.items()])
        kat, warna = kategori_nilai(score)
        st.markdown(f'<div style="background:white; border-left:5px solid {warna}; padding:2rem; border-radius:8px; text-align:center; margin-top:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);"><h3 style="margin:0; color:#475569;">Hasil Perkiraan Nilai</h3><h1 style="font-size:4rem; color:{warna}; margin:10px 0;">{score:.2f}</h1><h4 style="margin:0; color:{warna};">Kategori: {kat}</h4></div>', unsafe_allow_html=True)

elif st.session_state.current_page == "akun":
    st.title("Pengaturan Akun")
    with st.form("form_ganti_password"):
        old_pw = st.text_input("Password Lama", type="password")
        new_pw = st.text_input("Password Baru (min. 6 karakter)", type="password")
        confirm_pw = st.text_input("Ulangi Password Baru", type="password")
        if st.form_submit_button("Simpan Password Baru", type="primary"):
            if len(new_pw) < 6: st.error("Password minimal 6 karakter.")
            elif new_pw != confirm_pw: st.error("Password baru tidak cocok.")
            else:
                users = load_users()
                user_key = next((k for k, v in users.items() if v['nama'] == st.session_state.user_nama), None)
                if user_key and verify_password(old_pw, users[user_key]["password"]):
                    update_user_password(user_key, hash_password(new_pw))
                    st.success("Password berhasil diubah!")
                else: st.error("Password lama salah.")

elif st.session_state.current_page == "tentang":
    st.title("Panduan Penggunaan Aplikasi")
    st.markdown("""
    1. **Portal Siswa**: Siswa mendaftar mandiri (NISN & Password sama) dan mengisi survei.
    2. **Menu Nilai Raport**: Guru menginput nilai (bisa per mapel atau satu rata-rata).
    3. **Laporan Siswa**: Pilih siswa untuk melihat alasan nilai mereka naik/turun dan rekomendasi untuk guru.
    4. **Analisis Faktor**: Lihat alur hubungan faktor atau bandingkan pengaruh faktor per siswa berdasarkan model DAG.
    """)

elif st.session_state.current_page == "admin_users":
    st.title("Kelola Akun Guru & Admin")
    users = load_users()
    pending_users = {k: v for k, v in users.items() if v.get("status") == "pending"}
    if pending_users:
        st.markdown("### ⚠️ Menunggu Persetujuan")
        for u, data in pending_users.items():
            c1, c2, c3 = st.columns([3, 1, 1])
            c1.write(f"**{data['nama']}** ingin mendaftar sebagai `{u}`")
            if c2.button("✅ Setujui", key=f"app_{u}"): users[u]["status"] = "active"; save_users(users); st.rerun()
            if c3.button("❌ Tolak", key=f"rej_{u}"): del users[u]; save_users(users); st.rerun()
        st.markdown("---")
    with st.expander("➕ Tambah Pengguna Baru"):
        with st.form("form_tambah_user"):
            c1, c2 = st.columns(2)
            with c1:
                new_username = st.text_input("Username Baru")
                new_nama_user = st.text_input("Nama Lengkap")
            with c2:
                new_password = st.text_input("Password Awal", type="password")
                new_role = st.selectbox("Jabatan", ["guru", "admin"])
            if st.form_submit_button("Buat Akun", type="primary"):
                if new_username and new_nama_user and new_password:
                    if new_username in users: st.error("Username sudah digunakan.")
                    else:
                        users[new_username] = {"nama": new_nama_user, "password": hash_password(new_password), "role": new_role, "status": "active"}
                        save_users(users); st.success(f"Akun {new_nama_user} dibuat!"); st.rerun()
                else: st.error("Semua kolom wajib diisi.")
    st.markdown("---")
    st.subheader("Daftar Pengguna Aktif")
    for username, data in users.items():
        if data.get("status") == "pending": continue
        c1, c2, c3, c4 = st.columns([2, 2, 1, 2])
        c1.write(f"**{data['nama']}**"); c2.write(f"Username: `{username}`"); c3.write(f"Role: {data.get('role', 'guru').upper()}")
        with c4:
            if username not in SUPER_ADMIN and username != "admin":
                if st.button("Hapus Akun", key=f"del_user_{username}", type="primary"):
                    del users[username]; save_users(users); st.rerun()
            else: st.caption("Super Admin (Terkunci)")
        st.markdown("<hr style='margin: 5px 0; border-color: #e2e8f0;'>", unsafe_allow_html=True)