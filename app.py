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
import statsmodels.api as sm
import networkx as nx
import matplotlib.pyplot as plt
from dowhy import CausalModel
import warnings

warnings.filterwarnings("ignore")

from db_storage import (
    load_users, save_users, login_siswa, simpan_survei_siswa, 
    get_data_gabungan, update_nilai_siswa, tambah_siswa_baru, 
    update_user_password, daftar_user_baru, get_semua_akun_siswa, 
    import_data_historis_lengkap, reset_password_siswa, hapus_survei_siswa
)

SUPER_ADMIN = {"admin", "regina"}

# ================================================================
# ★ FUNGSI PARSER & KONSTANTA
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
    "Parental_Involvement": [
        "Orang tua menyediakan waktu untuk membantu saya belajar", 
        "Orang tua menghargai prestasi yang saya capai", 
        "Orang tua menyediakan fasilitas yang saya butuhkan untuk belajar", 
        "Orang tua memantau kegiatan belajar saya di rumah", 
        "Orang tua menanyakan tentang nilai atau tugas sekolah"
    ],
    "Academic_Anxiety": [
        "Saya merasa cemas menghadapi ujian akhir", 
        "Saya merasa tertekan dengan target nilai yang harus saya capai", 
        "Saya sulit memahami materi pelajaran"
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
    "Self_Efficacy": 4.634, "Academic_Anxiety": 3.547, "School_Support": 4.656, 
    "Parental_Involvement": 3.462, "Learning_Motivation": 4.292
}
RATA_RATA_NILAI = 83.780

# ================================================================
# ★ 4 MODEL DAG (Sesuai Jupyter Notebook)
# ================================================================
DAG_MODELS = {
    "Model 1: Baseline Direct Effects": {
        "deskripsi": "Model Input-Output / Black-Box. Semua konstruk mempengaruhi Prestasi secara langsung tanpa mediasi.",
        "edges": [
            ("Parental_Involvement", "Nilai Akademik"), ("School_Support", "Nilai Akademik"), 
            ("Self_Efficacy", "Nilai Akademik"), ("Learning_Motivation", "Nilai Akademik"), ("Academic_Anxiety", "Nilai Akademik")
        ]
    },
    "Model 2: Full Mediation (SDT)": {
        "deskripsi": "SDT (Ryan & Deci). Lingkungan sosial mempengaruhi prestasi melalui mekanisme psikologis.",
        "edges": [
            ("Parental_Involvement", "Learning_Motivation"), ("Parental_Involvement", "Self_Efficacy"), 
            ("Parental_Involvement", "Academic_Anxiety"), ("School_Support", "Learning_Motivation"), 
            ("School_Support", "Self_Efficacy"), ("School_Support", "Academic_Anxiety"),
            ("Learning_Motivation", "Nilai Akademik"), ("Self_Efficacy", "Nilai Akademik"), ("Academic_Anxiety", "Nilai Akademik")
        ]
    },
    "Model 3: Anxiety as Suppressor": {
        "deskripsi": "Lazarus & Folkman. Kecemasan menekan motivasi & self-efficacy, namun orang tua/sekolah juga punya efek langsung.",
        "edges": [
            ("Parental_Involvement", "Learning_Motivation"), ("Parental_Involvement", "Self_Efficacy"),
            ("Parental_Involvement", "Academic_Anxiety"), ("Parental_Involvement", "Nilai Akademik"),
            ("School_Support", "Learning_Motivation"), ("School_Support", "Self_Efficacy"),
            ("School_Support", "Academic_Anxiety"), ("School_Support", "Nilai Akademik"),
            ("Academic_Anxiety", "Learning_Motivation"), ("Academic_Anxiety", "Self_Efficacy"),
            ("Academic_Anxiety", "Nilai Akademik"), ("Learning_Motivation", "Nilai Akademik"),
            ("Self_Efficacy", "Nilai Akademik")
        ]
    },
    "Model 4: Full Structural (Ecological)": {
        "deskripsi": "Bronfenbrenner & Zimmerman. Model paling komprehensif dengan SE→LM, LM→AA, dan PI→SS.",
        "edges": [
            ("Parental_Involvement", "School_Support"), ("Parental_Involvement", "Learning_Motivation"),
            ("Parental_Involvement", "Self_Efficacy"), ("Parental_Involvement", "Academic_Anxiety"),
            ("Parental_Involvement", "Nilai Akademik"),
            ("School_Support", "Learning_Motivation"), ("School_Support", "Self_Efficacy"),
            ("School_Support", "Academic_Anxiety"), ("School_Support", "Nilai Akademik"),
            ("Self_Efficacy", "Learning_Motivation"), ("Self_Efficacy", "Nilai Akademik"),
            ("Learning_Motivation", "Academic_Anxiety"), ("Learning_Motivation", "Nilai Akademik"),
            ("Academic_Anxiety", "Nilai Akademik")
        ]
    }
}

# ================================================================
# HELPER
# ================================================================
def hash_password(p): return hashlib.sha256(p.encode()).hexdigest()
def verify_password(p, hashed): return hash_password(p) == hashed

def refresh_pengaruh(): 
    st.session_state.database_siswa = get_data_gabungan() or []

def kategori_nilai(n):
    if n >= 88: return "Sangat Baik", "#10B981"
    if n >= 84: return "Baik", "#2563EB"
    if n >= 80: return "Cukup", "#F0B900"
    return "Perlu Perhatian", "#EF5B67"

def run_scm_analysis(df_siswa, dag_edges):
    """Menjalankan DoWhy untuk menghitung ATE dan OLS untuk koefisien jalur."""
    if df_siswa.empty: 
        return {}, {}
    
    cols = ["Nilai Akademik", "Self_Efficacy", "Academic_Anxiety", "School_Support", "Parental_Involvement", "Learning_Motivation"]
    
    # Pastikan semua kolom ada dan konversi ke numeric
    for c in cols:
        if c not in df_siswa.columns:
            df_siswa[c] = 0.0
        df_siswa[c] = pd.to_numeric(df_siswa[c], errors='coerce').fillna(0.0)

    df = df_siswa[cols].copy()
    df = df[df["Nilai Akademik"] > 0] # Hanya siswa yang punya nilai
    df = df.dropna()
    
    if len(df) < 10: 
        return {}, {}

    # 1. Buat sintaks Graphviz untuk DoWhy
    graph_str = "digraph {\n" + "\n".join([f"{e[0]} -> {e[1]};" for e in dag_edges]) + "\n}"
    
    # 2. Hitung ATE dengan DoWhy
    ates = {}
    all_nodes = list(set([e[0] for e in dag_edges]))
    
    for trt in all_nodes:
        if trt == "Nilai Akademik": continue
        try:
            model = CausalModel(data=df, treatment=trt, outcome="Nilai Akademik", graph=graph_str)
            identified = model.identify_effect(proceed_when_unidentifiable=True)
            est = model.estimate_effect(identified, method_name="backdoor.linear_regression")
            ates[trt] = float(est.value)
        except:
            ates[trt] = 0.0
            
    # 3. Hitung OLS (untuk koefisien garis di DAG)
    ols_params = {}
    try:
        X = sm.add_constant(df[all_nodes])
        res = sm.OLS(df["Nilai Akademik"], X).fit()
        for node in all_nodes:
            ols_params[node] = float(res.params.get(node, 0))
    except:
        pass
        
    return ates, ols_params

def draw_dag_matplotlib(edges, ates):
    """Menggambar DAG menggunakan NetworkX & Matplotlib"""
    G = nx.DiGraph()
    G.add_edges_from([(e[0], e[1]) for e in edges])
    
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # Warna node
    node_colors = {
        "Self_Efficacy": "#9C27B0", "Academic_Anxiety": "#F44336",
        "School_Support": "#4CAF50", "Parental_Involvement": "#2196F3",
        "Learning_Motivation": "#FF9800", "Nilai Akademik": "#009688"
    }
    colors = [node_colors.get(node, "#888") for node in G.nodes()]
    
    try:
        pos = nx.nx_agraph.graphviz_layout(G, prog="dot")
    except:
        pos = nx.spring_layout(G, seed=42)
        
    nx.draw_networkx_nodes(G, pos, node_color=colors, node_size=3000, alpha=0.9, ax=ax)
    nx.draw_networkx_labels(G, pos, font_color="white", font_weight="bold", ax=ax)
    nx.draw_networkx_edges(G, pos, arrows=True, arrowstyle="-|>", arrowsize=20, width=2, edge_color="#555", ax=ax)
    
    # Tambahkan label ATE jika ada
    if ates:
        edge_labels = {}
        for e in edges:
            if e[0] in ates:
                edge_labels[e] = f"{ates[e[0]]:+.2f}"
        nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=9, ax=ax, bbox=dict(alpha=0.7, edgecolor='none'))
        
    ax.axis("off")
    st.pyplot(fig)
    plt.close(fig)

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
.dag-card { background: white; padding: 20px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
.dag-card h3 { margin: 0 0 5px 0; font-size: 1.1rem; color: #0f172a; }
.dag-card p { margin: 0 0 10px 0; font-size: 0.9rem; color: #475569; }
</style>
""", unsafe_allow_html=True)

# ================================================================
# STATE MANAGEMENT
# ================================================================
for key, val in {
    "logged_in": False, "user_nama": None, "user_role": None,
    "database_siswa": [], "current_page": "home", "last_menu": "Beranda",
    "siswa_login_nisn": None, "edit_nisn": None, "preview_nisn": None,
    "page_survei": 1, "page_nilai": 1, "custom_dags": {}
}.items():
    if key not in st.session_state: st.session_state[key] = val

if st.session_state.logged_in: refresh_pengaruh()

# ================================================================
# KUESIONER & LOGIN
# ================================================================
def render_kuesioner_siswa(nisn, is_preview=False):
    if is_preview:
        st.info(f"👁️ **Mode Pratinjau:** Anda melihat tampilan kuesioner siswa dengan NISN {nisn}.")
        if st.button("⬅️ Kembali ke Daftar Survei"):
            st.session_state.current_page = "list_survei"; st.session_state.preview_nisn = None; st.rerun()
    else: st.success(f"✅ Login berhasil! NISN: {nisn}")
        
    st.markdown("### 📝 Kuesioner Siswa")
    df = pd.DataFrame(st.session_state.database_siswa)
    siswa_data = df[df['NISN'] == str(nisn)].iloc[0] if not df.empty and str(nisn) in df['NISN'].values else {}
    
    with st.form("form_kuesioner"):
        st.markdown("#### 🅰️ Bagian A — Identitas & Latar Belakang")
        c_id1, c_id2 = st.columns(2)
        nama_input = c_id1.text_input("Nama Lengkap", value=siswa_data.get("Nama", "") if isinstance(siswa_data, pd.Series) else "")
        kelas_input = c_id2.text_input("Kelas", value=siswa_data.get("Kelas", "") if isinstance(siswa_data, pd.Series) else "")
        
        non_likert_data = {}
        c1, c2 = st.columns(2)
        for i, (nama_kolom, cfg) in enumerate(NON_LIKERT_ITEMS.items()):
            col = c1 if i % 2 == 0 else c2
            val = siswa_data.get(nama_kolom, cfg["default"]) if isinstance(siswa_data, pd.Series) else cfg["default"]
            with col:
                if cfg["type"] == "number":
                    clean_val = safe_int(val, cfg["default"]) if isinstance(cfg["default"], int) else safe_float(val, cfg["default"])
                    non_likert_data[nama_kolom] = st.number_input(cfg["label"], min_value=cfg["min"], max_value=cfg["max"], value=clean_val, step=cfg.get("step", 1), key=f"nl_{i}")
                else:
                    idx_val = 2
                    if pd.notna(val):
                        try: idx_val = cfg["options"].index(int(val))
                        except: idx_val = cfg["options"].index(safe_pendidikan(val))
                    non_likert_data[nama_kolom] = st.selectbox(cfg["label"], options=cfg["options"], index=idx_val, format_func=lambda x, L=cfg["labels"], O=cfg["options"]: L[O.index(x)], key=f"nl_{i}")
        
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
            if not nama_input or not kelas_input: st.error("Mohon isi Nama dan Kelas.")
            else:
                payload = {**non_likert_data, "Nama": nama_input, "Kelas": kelas_input}
                for construct, items in KONSTRUK_ITEMS.items():
                    vals = [construct_answers[i] for i in items]
                    if construct in REVERSE_ITEMS: vals = [6 - v for v in vals]
                    payload[construct] = sum(vals) / len(vals)
                simpan_survei_siswa(nisn, payload)
                st.balloons(); st.success("🎉 Kuesioner disimpan!")
                refresh_pengaruh()
                if is_preview: 
                    st.session_state.current_page = "list_survei"
                    st.session_state.last_menu = "Jawaban Survei"
                    st.session_state.preview_nisn = None
                    st.session_state.edit_nisn = None
                else: 
                    st.session_state.siswa_login_nisn = None
                st.rerun()

def halaman_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br><h1 style='text-align:center; color:var(--teal); font-size: 2.8rem;'>🎓 SAA Analytics</h1>", unsafe_allow_html=True)
        tab_login, tab_daftar, tab_siswa = st.tabs(["🔐 Login Guru/Admin", "📝 Daftar Akun Guru", "🎒 Portal Siswa"])
        with tab_login:
            with st.form("login_form"):
                username = st.text_input("Username"); password = st.text_input("Password", type="password")
                if st.form_submit_button("Masuk", use_container_width=True):
                    users = load_users()
                    if username in users:
                        if users[username].get("status") == "pending": st.warning("Akun menunggu persetujuan Admin.")
                        elif verify_password(password, users[username]["password"]):
                            st.session_state.logged_in = True; st.session_state.user_nama = users[username]["nama"]
                            st.session_state.user_role = users[username]["role"]; st.session_state.current_page = "home"
                            refresh_pengaruh(); st.rerun()
                        else: st.error("Password salah.")
                    else: st.error("Username tidak ditemukan.")
        with tab_daftar:
            with st.form("register_form"):
                reg_nama = st.text_input("Nama Lengkap"); reg_user = st.text_input("Username"); reg_pass = st.text_input("Password", type="password")
                if st.form_submit_button("Daftar", use_container_width=True):
                    sukses, pesan = daftar_user_baru(reg_user, reg_nama, hash_password(reg_pass), "guru")
                    if sukses: st.success(pesan)
                    else: st.error(pesan)
        with tab_siswa:
            if not st.session_state.siswa_login_nisn:
                with st.form("form_login_siswa"):
                    nisn_input = st.text_input("NISN"); pwd_input = st.text_input("Password (NISN)", type="password")
                    if st.form_submit_button("Masuk / Daftar", type="primary"):
                        sukses, pesan = login_siswa(nisn_input, pwd_input)
                        if sukses: st.session_state.siswa_login_nisn = nisn_input; st.rerun()
                        else: st.error(pesan)
            else: render_kuesioner_siswa(st.session_state.siswa_login_nisn, is_preview=False)

if not st.session_state.logged_in:
    if st.session_state.preview_nisn: render_kuesioner_siswa(st.session_state.preview_nisn, is_preview=True)
    else: halaman_login()
    st.stop()

# ================================================================
# SIDEBAR MENU
# ================================================================
with st.sidebar:
    st.markdown("""<div class="sb-logo"><div class="sb-logo-text"><h2>SAA<br><span style="font-size:1rem; font-weight:400;">Analytics</span></h2></div></div><hr>""", unsafe_allow_html=True)
    menu_options = ["Beranda", "Siswa", "Jawaban Survei", "Nilai Raport", "Laporan Siswa", "Analisis Faktor", "Simulasi Nilai", "Akun", "Panduan"]
    menu_icons = ["house-door", "people", "clipboard-data", "bar-chart", "file-earmark-text", "diagram-3", "magic", "person-circle", "info-circle"]
    if st.session_state.user_role == "admin": menu_options.extend(["Kelola Pengguna", "Keluar"]); menu_icons.extend(["gear", "box-arrow-right"])
    else: menu_options.append("Keluar"); menu_icons.append("box-arrow-right")

    mapping_menu = {"Beranda": "home", "Siswa": "monitoring", "Jawaban Survei": "list_survei", "Nilai Raport": "list_nilai", "Laporan Siswa": "rangkuman", "Analisis Faktor": "kausal", "Simulasi Nilai": "prediksi", "Akun": "akun", "Panduan": "tentang", "Kelola Pengguna": "admin_users", "Keluar": "logout"}
    current_page_mapped = st.session_state.current_page
    if current_page_mapped == "edit_x": current_page_mapped = "list_survei"
    if current_page_mapped == "edit_y": current_page_mapped = "list_nilai"

    default_idx = next((i for i, opt in enumerate(menu_options) if mapping_menu.get(opt) == current_page_mapped), 0)
    menu_choice = option_menu(None, menu_options, icons=menu_icons, default_index=default_idx, key="sidebar_menu",
        styles={"nav-link-selected": {"background-color": "rgba(255,255,255,0.15)", "color": "#facc15 !important", "border-left": "4px solid #facc15"}})

    if menu_choice != st.session_state.last_menu:
        st.session_state.last_menu = menu_choice; st.session_state.page_survei = 1; st.session_state.page_nilai = 1
        if menu_choice == "Keluar":
            for k in ("logged_in", "preview_nisn", "siswa_login_nisn", "edit_nisn"):
                st.session_state[k] = None if k.endswith("_nisn") or k.startswith("preview") else False
            st.session_state.database_siswa = []
            st.rerun()
        else: st.session_state.current_page = mapping_menu.get(menu_choice, "home")

    st.markdown(f"<br><br><hr><div style='padding-left: 10px;'><p style='color: #94a3b8; font-size: 0.8rem;'>Masuk sebagai:</p><p style='color: white; font-weight: bold;'>👤 {st.session_state.user_nama}</p></div>", unsafe_allow_html=True)

# ================================================================
# HALAMAN: BERANDA
# ================================================================
if st.session_state.current_page == "home":
    st.markdown(f"""<div class="top-header" style="margin-bottom: 2rem;"><h1>Halo, {st.session_state.user_nama}</h1></div>""", unsafe_allow_html=True)
    refresh_pengaruh(); df = pd.DataFrame(st.session_state.database_siswa)
    total_siswa = len(df)
    
    nilai_num = pd.to_numeric(df.get("Nilai Akademik", pd.Series(dtype=float)), errors="coerce").fillna(0)
    n_terisi = int((nilai_num > 0).sum()) if not df.empty else 0
    survei_lengkap = df["Self_Efficacy"].notna().sum() if "Self_Efficacy" in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="kpi-card"><h2>{total_siswa}</h2><p>Siswa</p></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><h2>37</h2><p>Jumlah Pertanyaan</p></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><h2>{survei_lengkap} siswa</h2><p>Survei Lengkap</p></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card"><h2>{n_terisi} / {total_siswa}</h2><p>Nilai Terisi</p></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_chart, col_checklist = st.columns([2, 1])
    with col_chart:
        st.markdown("### Sebaran Nilai Siswa")
        if not df.empty and 'Nilai Akademik' in df.columns and (nilai_num > 0).any():
            df_valid = df[nilai_num > 0]
            fig = px.histogram(df_valid, x="Nilai Akademik", nbins=10, color_discrete_sequence=["#0d9488"])
            fig.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", xaxis_title="Rentang Nilai", yaxis_title="Jumlah Siswa")
            st.plotly_chart(fig, use_container_width=True)
            st.caption(f"Rata-rata nilai sekolah: **{df_valid['Nilai Akademik'].mean():.2f}**")
        else: st.info("Belum ada data nilai.")
            
    with col_checklist:
        st.markdown("""<div style="background: white; padding: 20px; border-radius: 8px; border: 1px solid #e2e8f0; height: 100%;"><h4 style="margin-top:0; font-size:1.1rem; color: #0f172a;">Data yang Belum Lengkap</h4>""", unsafe_allow_html=True)
        x_belum_isi = total_siswa - survei_lengkap
        st.markdown(f"""<ul style="list-style-type: none; padding-left: 0; font-size: 0.9rem; line-height: 2; color: #475569;"><li>📄 <b>{x_belum_isi} siswa</b> belum mengisi survei</li></ul><hr></div>""", unsafe_allow_html=True)

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
                new_nisn = st.text_input("NISN"); new_nama = st.text_input("Nama Lengkap")
            with c2:
                new_kelas = st.text_input("Kelas"); new_nilai = st.number_input("Nilai Raport", min_value=0.0, max_value=100.0)
            if st.form_submit_button("Simpan", type="primary"):
                if new_nisn and new_nama:
                    tambah_siswa_baru(new_nisn, new_nama, new_kelas, new_nilai)
                    st.success("Berhasil!"); refresh_pengaruh(); st.rerun()

    with tab_import:
        uploaded = st.file_uploader("Upload File Excel (.xlsx)", type=["xlsx"])
        if uploaded:
            try:
                df_upload = pd.read_excel(uploaded)
                st.dataframe(df_upload.head())
                if st.button("Simpan Nilai ke Database", type="primary"):
                    count = import_data_historis_lengkap(df_upload)
                    st.success(f"✅ {count} data dimasukkan!"); refresh_pengaruh(); st.rerun()
            except Exception as e: st.error(f"Error: {e}")

    with tab_daftar:
        df_siswa = pd.DataFrame(st.session_state.database_siswa)
        if not df_siswa.empty:
            for idx, row in df_siswa.iterrows():
                c1, c2, c3, c4, c5, c6 = st.columns([1.5, 3, 2, 1, 2, 1.5])
                c1.write(f"**{row.get('NISN', '-')}**"); c2.write(f"**{row['Nama']}**"); c3.write(f"Kelas: {row.get('Kelas', '-')}"); c5.write(f"Nilai: {row.get('Nilai Akademik', '0')}")
                with c6:
                    b1, b2 = st.columns(2)
                    if b1.button("👁️ Portal", key=f"prev_{row['NISN']}"): 
                        st.session_state.edit_nisn = row['NISN']
                        st.session_state.current_page = "edit_x"
                        st.session_state.last_menu = "Jawaban Survei"
                        st.rerun()
                    if b2.button("✏️ Edit", key=f"edit_{row['NISN']}"): 
                        st.session_state.edit_nisn = row['NISN']; st.session_state.current_page = "edit_x"; st.rerun()
                st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)

# ================================================================
# HALAMAN: DAFTAR SURVEI (PAGINATION)
# ================================================================
elif st.session_state.current_page == "list_survei":
    st.title("Data Survei Siswa")
    refresh_pengaruh(); df_siswa = pd.DataFrame(st.session_state.database_siswa)
    if not df_siswa.empty:
        cari = st.text_input("🔍 Cari nama atau NISN:")
        if cari: df_siswa = df_siswa[df_siswa['Nama'].astype(str).str.contains(cari, case=False, na=False) | df_siswa['NISN'].astype(str).str.contains(cari, case=False, na=False)]
        
        items_per_page = 10; total_pages = math.ceil(len(df_siswa) / items_per_page)
        if st.session_state.page_survei > total_pages: st.session_state.page_survei = total_pages
        if st.session_state.page_survei < 1: st.session_state.page_survei = 1
            
        df_page = df_siswa.iloc[(st.session_state.page_survei - 1) * items_per_page : (st.session_state.page_survei) * items_per_page]

        for idx, row in df_page.iterrows():
            c1, c2, c3, c4, c5, c6 = st.columns([1.5, 4, 3, 1.2, 1.2, 1.2])
            c1.write(f"**{row['NISN']}**"); c2.write(f"**{row['Nama']}**")
            status = "Lengkap ✅" if pd.notna(row.get("Self_Efficacy")) else "Belum ❌"
            c3.write(f"Terisi: {status}")
            with c4:
                if st.button("Ubah", key=f"edit_survei_{row['NISN']}"): 
                    st.session_state.edit_nisn = row['NISN']; st.session_state.current_page = "edit_x"; st.rerun()
            with c5:
                if st.button("🔑 Reset Pass", key=f"reset_pw_{row['NISN']}"): reset_password_siswa(row['NISN']); st.success("Direset!"); st.rerun()
            with c6:
                if st.button("🗑️ Hapus", key=f"del_survei_{row['NISN']}", type="primary"): hapus_survei_siswa(row['NISN']); refresh_pengaruh(); st.rerun()
            st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)

        # --- PAGINATION NAVIGATION ---
        p1, p2, p3 = st.columns([1, 2, 1])
        with p1:
            if st.button("⬅️ Sebelumnya", disabled=(st.session_state.page_survei == 1), use_container_width=True): st.session_state.page_survei -= 1; st.rerun()
        with p2:
            c_l, c_m, c_r = st.columns([1, 1.5, 1])
            with c_m:
                col_t1, col_sel, col_t2 = st.columns([1, 1, 1])
                with col_t1: st.markdown("<div style='text-align:right; padding-top:12px;'>Halaman</div>", unsafe_allow_html=True)
                with col_sel:
                    sel_page = st.selectbox("Lompat", options=list(range(1, total_pages + 1)), index=st.session_state.page_survei - 1, label_visibility="collapsed")
                    if sel_page != st.session_state.page_survei: st.session_state.page_survei = sel_page; st.rerun()
                with col_t2: st.markdown(f"<div style='text-align:left; padding-top:12px;'>dari <b>{total_pages}</b></div>", unsafe_allow_html=True)
        with p3:
            if st.button("Berikutnya ➡️", disabled=(st.session_state.page_survei == total_pages), use_container_width=True): st.session_state.page_survei += 1; st.rerun()

# ================================================================
# HALAMAN: DAFTAR NILAI (PAGINATION)
# ================================================================
elif st.session_state.current_page == "list_nilai":
    st.title("Data Nilai Raport Siswa")
    df_siswa = pd.DataFrame(st.session_state.database_siswa)
    if not df_siswa.empty:
        cari = st.text_input("🔍 Cari nama atau NISN:")
        if cari: df_siswa = df_siswa[df_siswa['Nama'].astype(str).str.contains(cari, case=False, na=False) | df_siswa['NISN'].astype(str).str.contains(cari, case=False, na=False)]
        
        items_per_page = 10; total_pages = math.ceil(len(df_siswa) / items_per_page)
        if st.session_state.page_nilai > total_pages: st.session_state.page_nilai = total_pages
        if st.session_state.page_nilai < 1: st.session_state.page_nilai = 1
            
        df_page = df_siswa.iloc[(st.session_state.page_nilai - 1) * items_per_page : (st.session_state.page_nilai) * items_per_page]

        for idx, row in df_page.iterrows():
            c1, c2, c3, c4, c5 = st.columns([1.5, 4, 2, 2, 1])
            c1.write(f"**{row['NISN']}**"); c2.write(f"**{row['Nama']}**"); c3.write(f"Nilai: **{safe_float(row.get('Nilai Akademik', 0)):.2f}**"); c4.write(f"Status: {'Lengkap' if safe_float(row.get('Nilai Akademik',0))>0 else 'Kosong'}")
            with c5:
                if st.button("Ubah", key=f"edit_nilai_{row['NISN']}"): st.session_state.edit_nisn = row['NISN']; st.session_state.current_page = "edit_y"; st.rerun()
            st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)

        # --- PAGINATION NAVIGATION ---
        p1, p2, p3 = st.columns([1, 2, 1])
        with p1:
            if st.button("⬅️ Sebelumnya", disabled=(st.session_state.page_nilai == 1), use_container_width=True, key="prev_nilai"): st.session_state.page_nilai -= 1; st.rerun()
        with p2:
            c_l, c_m, c_r = st.columns([1, 1.5, 1])
            with c_m:
                col_t1, col_sel, col_t2 = st.columns([1, 1, 1])
                with col_t1: st.markdown("<div style='text-align:right; padding-top:12px;'>Halaman</div>", unsafe_allow_html=True)
                with col_sel:
                    sel_page = st.selectbox("Lompat", options=list(range(1, total_pages + 1)), index=st.session_state.page_nilai - 1, label_visibility="collapsed", key="sel_nilai")
                    if sel_page != st.session_state.page_nilai: st.session_state.page_nilai = sel_page; st.rerun()
                with col_t2: st.markdown(f"<div style='text-align:left; padding-top:12px;'>dari <b>{total_pages}</b></div>", unsafe_allow_html=True)
        with p3:
            if st.button("Berikutnya ➡️", disabled=(st.session_state.page_nilai == total_pages), use_container_width=True, key="next_nilai"): st.session_state.page_nilai += 1; st.rerun()

# ================================================================
# HALAMAN: EDIT SURVEI & NILAI
# ================================================================
elif st.session_state.current_page == "edit_x":
    df = pd.DataFrame(st.session_state.database_siswa)
    if st.session_state.get("edit_nisn") is None or df.empty:
        st.warning("Pilih siswa dulu."); 
        if st.button("Kembali"): st.session_state.current_page = "list_survei"; st.rerun()
    else:
        if st.button("⬅️ Kembali ke Daftar Survei"): st.session_state.current_page = "list_survei"; st.session_state.last_menu = "Jawaban Survei"; st.rerun()
        siswa = df[df['NISN'] == str(st.session_state.edit_nisn)].iloc[0]
        render_kuesioner_siswa(siswa['NISN'], is_preview=True)

elif st.session_state.current_page == "edit_y":
    df = pd.DataFrame(st.session_state.database_siswa)
    if st.session_state.get("edit_nisn") is None or df.empty:
        st.warning("Pilih siswa dulu.")
        if st.button("Kembali"): st.session_state.current_page = "list_nilai"; st.rerun()
    else:
        if st.button("⬅️ Kembali ke Daftar Nilai"): st.session_state.current_page = "list_nilai"; st.session_state.last_menu = "Nilai Raport"; st.rerun()
        siswa = df[df['NISN'] == str(st.session_state.edit_nisn)].iloc[0]
        nilai_skrg = safe_float(siswa.get('Nilai Akademik', 0))
        
        with st.form("form_nilai_rata"):
            nilai_tunggal = st.number_input("Masukkan Nilai Raport", min_value=0.0, max_value=100.0, value=nilai_skrg if nilai_skrg > 0 else 80.0, step=0.1)
            if st.form_submit_button("Simpan data Y", type="primary"):
                sukses, pesan = update_nilai_siswa(siswa['NISN'], {"Nilai": nilai_tunggal})
                if sukses: st.success(pesan); refresh_pengaruh(); st.session_state.current_page = "list_nilai"; st.rerun()

# ================================================================
# HALAMAN: LAPORAN SISWA (DINAMIS BERDASARKAN DAG)
# ================================================================
elif st.session_state.current_page == "rangkuman":
    st.title("Laporan Hasil Analisis Siswa")
    refresh_pengaruh(); df = pd.DataFrame(st.session_state.database_siswa)
    
    if not df.empty:
        # Pilihan DAG untuk perhitungan kontribusi
        dag_pilihan = st.selectbox("Pilih Model Analisis Faktor (DAG) untuk Laporan:", list(DAG_MODELS.keys()))
        
        df["Label"] = df["Nama"].astype(str) + " (" + df["Kelas"].astype(str).fillna("") + ")"
        siswa_terpilih = st.selectbox("Pilih Siswa:", df["Label"].tolist())
        
        if siswa_terpilih:
            siswa = df[df["Label"] == siswa_terpilih].iloc[0]
            nilai = safe_float(siswa.get("Nilai Akademik", 0)); kat, warna = kategori_nilai(nilai)
            
            st.markdown(f"### {siswa['Nama']} | Nilai: **{nilai:.2f}** ({kat})")
            
            with st.spinner("Menghitung kausalitas dinamis..."):
                # Hitung ATE secara dinamis berdasarkan DAG yang dipilih
                ates, _ = run_scm_analysis(df, DAG_MODELS[dag_pilihan]["edges"])
            
            nama_awam = {"Self_Efficacy": "Percaya Diri", "Academic_Anxiety": "Rasa Cemas", "School_Support": "Dukungan Sekolah", "Parental_Involvement": "Dukungan Orang Tua", "Learning_Motivation": "Semangat Belajar"}
            
            kontribusi = []
            for f, nama in nama_awam.items():
                ate = ates.get(f, 0.0)
                val = safe_float(siswa.get(f, 3.0)); base = BASELINE_ASPEK.get(f, 3.0)
                kontribusi.append({"Faktor": nama, "Pengaruh (Poin)": ate * (val - base)})
                
            df_kontrib = pd.DataFrame(kontribusi).sort_values("Pengaruh (Poin)")
            fig = px.bar(df_kontrib, x="Pengaruh (Poin)", y="Faktor", orientation="h", color="Pengaruh (Poin)", color_continuous_scale="RdBu", text=[f"{v:+.2f}" for v in df_kontrib["Pengaruh (Poin)"]])
            fig.update_traces(textposition="outside"); fig.update_layout(height=300, coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ================================================================
# HALAMAN: ANALISIS FAKTOR (CUSTOM DAG & DOWHY)
# ================================================================
elif st.session_state.current_page == "kausal":
    st.title("Analisis Faktor Nilai Siswa")
    st.write("Konstruk dibentuk dari fitur survei X (rata-rata butir), outcome = Nilai Akademik (Y). Setiap DAG dianalisis otomatis: ATE (penyesuaian backdoor) dan OLS.")
    
    tab_daftar, tab_create = st.tabs(["📂 Daftar DAG & Analisis", "➕ Create DAG"])
    
    # TAB 1: DAFTAR DAG
    with tab_daftar:
        st.markdown("### Pilih DAG untuk Dianalisis")
        all_models = list(DAG_MODELS.keys()) + list(st.session_state.custom_dags.keys())
        
        if all_models:
            pilihan_lihat = st.selectbox("Model DAG:", all_models)
            if pilihan_lihat in DAG_MODELS: info = DAG_MODELS[pilihan_lihat]
            else: info = st.session_state.custom_dags[pilihan_lihat]
            
            st.info(info["deskripsi"])
            
            if st.button("🚀 Jalankan Analisis DoWhy & OLS", type="primary"):
                refresh_pengaruh(); df = pd.DataFrame(st.session_state.database_siswa)
                
                if df.empty or len(df[df["Nilai Akademik"] > 0]) < 10:
                    st.warning("Data siswa dengan nilai kurang dari 10. Analisis membutuhkan minimal 10 data valid.")
                else:
                    with st.spinner("Menghitung Average Treatment Effect (ATE) dan Regresi..."):
                        ates, ols_params = run_scm_analysis(df, info["edges"])
                        
                        st.markdown("#### Hasil Average Treatment Effect (DoWhy)")
                        cols_ate = st.columns(len(ates))
                        for idx, (k, v) in enumerate(ates.items()):
                            with cols_ate[idx]:
                                st.metric(label=k.replace('_', ' '), value=f"{v:+.3f}")
                                
                        st.markdown("#### Visualisasi Diagram DAG (dengan nilai ATE)")
                        draw_dag_matplotlib(info["edges"], ates)

    # TAB 2: CREATE DAG
    with tab_create:
        st.info("Buat DAG kustom. Model akan ditambahkan ke daftar dan dapat dianalisis secara otomatis.")
        
        with st.form("form_create_dag"):
            nama_dag = st.text_input("Nama DAG", placeholder="mis. Model lingkungan → psikologis → prestasi")
            st.markdown("### Pilih Node yang Digunakan (Maks 6)")
            
            cols = st.columns(3); nodes_input = []
            defaults = ["Self_Efficacy", "Academic_Anxiety", "School_Support", "Parental_Involvement", "Learning_Motivation", "Nilai Akademik"]
            for i in range(6):
                with cols[i % 3]:
                    val = st.text_input(f"Node {i+1}", value=defaults[i], key=f"node_{i}")
                    if val.strip(): nodes_input.append(val.strip())
            
            if st.form_submit_button("Create DAG (Direct Effects Only)", type="primary"):
                if not nama_dag or len(nodes_input) < 1: 
                    st.error("Isi Nama DAG dan minimal satu Node.")
                else:
                    # Asumsi default: semua node kecuali 'Nilai Akademik' mempengaruhi 'Nilai Akademik'
                    new_edges = [(node, "Nilai Akademik") for node in nodes_input if node != "Nilai Akademik"]
                    st.session_state.custom_dags[nama_dag] = {
                        "deskripsi": f"Model kustom dibuat: {datetime.now().strftime('%Y-%m-%d')}.",
                        "edges": new_edges
                    }
                    st.success(f"DAG '{nama_dag}' berhasil dibuat! Silakan cek tab Daftar DAG."); st.rerun()

# ================================================================
# HALAMAN LAINNYA
# ================================================================
elif st.session_state.current_page == "prediksi":
    st.title("Simulasi Nilai Siswa")
    st.write("Geser nilai untuk melihat perkiraan nilai (Berdasarkan Model 1).")
    c1, c2 = st.columns(2)
    with c1:
        p_se  = st.slider("Percaya Diri", 1.0, 5.0, BASELINE_ASPEK["Self_Efficacy"], 0.1)
        p_mot = st.slider("Semangat Belajar", 1.0, 5.0, BASELINE_ASPEK["Learning_Motivation"], 0.1)
        p_cem = st.slider("Rasa Cemas", 1.0, 5.0, BASELINE_ASPEK["Academic_Anxiety"], 0.1)
    with c2:
        p_ortu = st.slider("Dukungan Orang Tua", 1.0, 5.0, BASELINE_ASPEK["Parental_Involvement"], 0.1)
        p_sekolah = st.slider("Dukungan Sekolah", 1.0, 5.0, BASELINE_ASPEK["School_Support"], 0.1)
        
    if st.button("Hitung Perkiraan Nilai", type="primary"):
        profil = {"Self_Efficacy": p_se, "Learning_Motivation": p_mot, "Academic_Anxiety": p_cem, "Parental_Involvement": p_ortu, "School_Support": p_sekolah}
        # Gunakan hardcode dari Jupyter Notebook (Model 1)
        ate_model1 = {"Parental_Involvement": 1.4212, "School_Support": -1.4545, "Self_Efficacy": 2.9292, "Learning_Motivation": 0.2533, "Academic_Anxiety": 1.2613}
        score = RATA_RATA_NILAI + sum([ate_model1[k] * (v - BASELINE_ASPEK[k]) for k, v in profil.items()])
        kat, warna = kategori_nilai(score)
        st.markdown(f'<div style="background:white; border-left:5px solid {warna}; padding:2rem; text-align:center; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);"><h3>Hasil Perkiraan Nilai</h3><h1 style="font-size:4rem; color:{warna}; margin:10px 0;">{score:.2f}</h1><h4>Kategori: {kat}</h4></div>', unsafe_allow_html=True)

elif st.session_state.current_page == "akun":
    st.title("Pengaturan Akun")
    with st.form("form_ganti_password"):
        old_pw = st.text_input("Password Lama", type="password"); new_pw = st.text_input("Password Baru", type="password"); confirm_pw = st.text_input("Ulangi Password Baru", type="password")
        if st.form_submit_button("Simpan", type="primary"):
            if len(new_pw) < 6: st.error("Minimal 6 karakter.")
            elif new_pw != confirm_pw: st.error("Tidak cocok.")
            else:
                users = load_users(); user_key = next((k for k, v in users.items() if v['nama'] == st.session_state.user_nama), None)
                if user_key and verify_password(old_pw, users[user_key]["password"]):
                    update_user_password(user_key, hash_password(new_pw)); st.success("Berhasil diubah!")
                else: st.error("Password lama salah.")

elif st.session_state.current_page == "tentang":
    st.title("Panduan Penggunaan Aplikasi")
    st.markdown("1. **Portal Siswa**: Siswa mendaftar mandiri dan mengisi survei.\n2. **Menu Nilai Raport**: Guru menginput nilai.\n3. **Laporan Siswa**: Melihat alasan nilai naik/turun berdasarkan model kausal terpilih.\n4. **Analisis Faktor**: Menjalankan simulasi DoWhy & melihat efek kausal (ATE).")

elif st.session_state.current_page == "admin_users":
    st.title("Kelola Akun")
    users = load_users(); pending_users = {k: v for k, v in users.items() if v.get("status") == "pending"}
    if pending_users:
        st.markdown("### ⚠️ Menunggu Persetujuan")
        for u, data in pending_users.items():
            c1, c2, c3 = st.columns([3, 1, 1]); c1.write(f"**{data['nama']}** (`{u}`)")
            if c2.button("✅ Setujui", key=f"app_{u}"): users[u]["status"] = "active"; save_users(users); st.rerun()
            if c3.button("❌ Tolak", key=f"rej_{u}"): del users[u]; save_users(users); st.rerun()
        st.markdown("---")
    
    with st.expander("➕ Tambah Pengguna Baru"):
        with st.form("form_tambah_user"):
            c1, c2 = st.columns(2)
            with c1: new_username = st.text_input("Username Baru"); new_nama_user = st.text_input("Nama Lengkap")
            with c2: new_password = st.text_input("Password Awal", type="password"); new_role = st.selectbox("Jabatan", ["guru", "admin"])
            if st.form_submit_button("Buat Akun", type="primary"):
                if new_username and new_nama_user and new_password:
                    if new_username in users: st.error("Username sudah ada.")
                    else:
                        users[new_username] = {"nama": new_nama_user, "password": hash_password(new_password), "role": new_role, "status": "active"}
                        save_users(users); st.success("Akun dibuat!"); st.rerun()
                else: st.error("Isi semua kolom.")
    
    st.markdown("---")
    st.subheader("Daftar Pengguna Aktif")
    for username, data in users.items():
        if data.get("status") == "pending": continue
        c1, c2, c3, c4 = st.columns([2, 2, 1, 2])
        c1.write(f"**{data['nama']}**"); c2.write(f"`{username}`"); c3.write(f"{data.get('role', 'guru').upper()}")
        with c4:
            if username not in SUPER_ADMIN and username != "admin":
                if st.button("Hapus Akun", key=f"del_user_{username}", type="primary"): del users[username]; save_users(users); st.rerun()
            else: st.caption("Super Admin (Terkunci)")
        st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)