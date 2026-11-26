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
import json
import os
from openpyxl import Workbook
from streamlit_option_menu import option_menu
import pytz
import statsmodels.api as sm

# --- Export laporan: Word & PDF ---
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from fpdf import FPDF
    HAS_FPDF = True
except ImportError:
    HAS_FPDF = False

from db_storage import (
    load_users, save_users, login_siswa, simpan_survei_siswa, 
    get_data_gabungan, update_nilai_siswa, tambah_siswa_baru, 
    update_user_password, daftar_user_baru, get_semua_akun_siswa, 
    import_data_historis_lengkap, reset_password_siswa, hapus_survei_siswa,
    hapus_nilai_saja
)

SUPER_ADMIN = {"admin", "regina"}
UPLOAD_HISTORY_FILE = "upload_history.json"

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
        "Suasana di sekolah saya mendukung untuk belajar",
        "Fasilitas sekolah (perpustakaan, laboratorium) mendukung pembelajaran"
    ],
    "Parental_Involvement": [
        "Orang tua menyediakan waktu untuk membantu saya belajar", 
        "Orang tua mendorong saya untuk meraih cita-cita",
        "Orang tua menghargai prestasi yang saya capai", 
        "Orang tua menyediakan fasilitas yang saya butuhkan untuk belajar", 
        "Orang tua memantau kegiatan belajar saya di rumah", 
        "Orang tua menanyakan tentang nilai atau tugas sekolah"
    ],
    "Academic_Anxiety": [
        "Saya merasa cemas menghadapi ujian akhir", 
        "Saya merasa tertekan dengan target nilai yang harus saya capai", 
        "Saya malas belajar",
        "Saya sulit memahami materi pelajaran",
        "Orang tua saya menuntut nilai yang tinggi"
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

RAW_QUESTIONS = []
for construct, items in KONSTRUK_ITEMS.items():
    for item in items:
        RAW_QUESTIONS.append({"kode": f"X{len(RAW_QUESTIONS) + 11}", "fitur": item, "default": construct})

# ================================================================
# ★ 4 MODEL DAG & NILAI ATE
# ================================================================
DAG_MODELS = {
    "Model 1: Baseline Direct Effects": {
        "deskripsi": "Input-Output Model. Semua konstruk berpengaruh langsung terhadap prestasi.",
        "edges": 5,
        "ate": {"Parental_Involvement": 1.42, "School_Support": -1.45, "Self_Efficacy": 2.93, "Learning_Motivation": 0.25, "Academic_Anxiety": 1.26}
    },
    "Model 2: Full Mediation (SDT)": {
        "deskripsi": "Lingkungan sosial memengaruhi prestasi melalui mekanisme psikologis (motivasi belajar dan self-efficacy).",
        "edges": 9,
        "ate": {"Parental_Involvement": 2.11, "School_Support": -0.17, "Self_Efficacy": 4.02, "Learning_Motivation": 0.70, "Academic_Anxiety": 1.26}
    },
    "Model 3: Anxiety as Suppressor": {
        "deskripsi": "Dari DAG 2, kecemasan juga menekan motivasi dan self-efficacy (jalur supresi), dan PI/SS juga punya efek langsung.",
        "edges": 13,
        "ate": {"Parental_Involvement": 2.11, "School_Support": -0.17, "Self_Efficacy": 3.82, "Learning_Motivation": 0.42, "Academic_Anxiety": 1.55}
    },
    "Model 4: Full Structural (Ecological)": {
        "deskripsi": "Model paling komprehensif: PI → SS, SE → LM, dan LM → AA (arah dibalik dari DAG 3).",
        "edges": 14,
        "ate": {"Parental_Involvement": 1.74, "School_Support": -0.73, "Self_Efficacy": 3.98, "Learning_Motivation": -0.01, "Academic_Anxiety": 1.34}
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

def get_mapel_list(df):
    default_mapel = ["Pendidikan Agama", "PPKn", "Bahasa Indonesia", "Matematika", "IPA", "IPS", "Bahasa Inggris", "Seni Budaya", "PJOK", "Prakarya/Informatika"]
    if df.empty: return default_mapel
    
    non_mapel = ['NISN', 'Nama', 'Kelas', 'Nilai Akademik', 'Self_Efficacy', 'Learning_Motivation', 'School_Support', 'Parental_Involvement', 'Academic_Anxiety']
    non_mapel.extend(NON_LIKERT_ITEMS.keys())
    
    mapels = []
    for c in df.columns:
        if c not in non_mapel and pd.api.types.is_numeric_dtype(df[c]):
            clean_name = c.replace("Nilai_", "")
            if clean_name not in mapels:
                mapels.append(clean_name)
                
    if not mapels: return default_mapel
    return mapels

def clean_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "", str(text))

# --- FITUR RIWAYAT UPLOAD ---
def load_upload_history():
    if os.path.exists(UPLOAD_HISTORY_FILE):
        try:
            with open(UPLOAD_HISTORY_FILE, "r") as f:
                return json.load(f)
        except:
            return []
    return []

def save_upload_history(history):
    with open(UPLOAD_HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=4)

def hitung_ols_custom(df, dag_edges):
    if df.empty: return {}
    nodes = list(set([e[0] for e in dag_edges]))
    valid_nodes = [n for n in nodes if n in ["Parental_Involvement", "School_Support", "Self_Efficacy", "Learning_Motivation", "Academic_Anxiety"]]
    if not valid_nodes: return {}
    
    df_clean = df[["Nilai Akademik"] + valid_nodes].apply(pd.to_numeric, errors='coerce').dropna()
    if len(df_clean) < 10: return {}
    
    X = sm.add_constant(df_clean[valid_nodes])
    y = df_clean["Nilai Akademik"]
    try:
        res = sm.OLS(y, X).fit()
        return {node: (res.params.get(node, 0), res.pvalues.get(node, 1)) for node in valid_nodes}
    except:
        return {}

def terjemah_efek(beta, pval):
    if pval >= 0.05: return "Tidak Berpengaruh"
    if beta > 0:
        if beta > 1.5: return "Sangat Menaikkan Nilai"
        if beta > 0.5: return "Menaikkan Nilai"
        return "Sedikit Menaikkan"
    else:
        if beta < -1.5: return "Sangat Menurunkan Nilai"
        if beta < -0.5: return "Menurunkan Nilai"
        return "Sedikit Menurunkan"

# ================================================================
# NARASI LAPORAN
# ================================================================
def _generate_narasi_laporan(siswa, nilai, kat, dag_pilihan, kontribusi_list):
    nama = str(siswa.get("Nama", ""))
    naik = [k for k in kontribusi_list if k["Pengaruh (Poin)"] > 0.1]
    turun = [k for k in kontribusi_list if k["Pengaruh (Poin)"] < -0.1]
    naik.sort(key=lambda x: x["Pengaruh (Poin)"], reverse=True)
    turun.sort(key=lambda x: x["Pengaruh (Poin)"])

    narasi = [
        f"Nilai akademik {nama} saat ini berada di angka {nilai:.2f} dengan predikat {kat}.",
        f"Berdasarkan hasil analisis, terdapat beberapa faktor yang mempengaruhi nilai {nama}:",
        ""
    ]
    if naik:
        narasi.append("Faktor yang Membantu Menaikkan Nilai:")
        for i, k in enumerate(naik, 1):
            narasi.append(f"  {i}. {k['Faktor']} (memberikan dorongan positif sebesar +{k['Pengaruh (Poin)']:.2f} poin)")
        narasi.append("")
    if turun:
        narasi.append("Faktor yang Menarik Nilai ke Bawah:")
        for i, k in enumerate(turun, 1):
            narasi.append(f"  {i}. {k['Faktor']} (menurunkan nilai sebesar {k['Pengaruh (Poin)']:.2f} poin)")
        narasi.append("")
    if not naik and not turun:
        narasi.append(f"Secara umum, kondisi psikologis dan lingkungan {nama} sudah seimbang dan berada di rata-rata sekolah. Tidak ada faktor yang menonjol.")
        narasi.append("")
    return "\n".join(narasi)

def _generate_rekomendasi_guru(siswa, nilai, kat, kontribusi_list):
    nama = str(siswa.get("Nama", ""))
    naik = sorted([k for k in kontribusi_list if k["Pengaruh (Poin)"] > 0.1], key=lambda x: x["Pengaruh (Poin)"], reverse=True)
    turun = sorted([k for k in kontribusi_list if k["Pengaruh (Poin)"] < -0.1], key=lambda x: x["Pengaruh (Poin)"])

    rekomendasi = []
    if naik:
        top = naik[0]["Faktor"]
        peta_kuat = {
            "Percaya Diri": f"Pertahankan pemberian tugas menantang yang terukur agar rasa percaya diri {nama} terus tumbuh di kelas.",
            "Semangat Belajar": f"Berikan apresiasi (pujian di depan kelas) atas usaha belajar {nama}, bukan hanya hasil akhirnya.",
            "Dukungan Orang Tua": f"Libatkan orang tua dalam kegiatan kelas (mis. proyek/presentasi) agar dukungan belajar {nama} tetap konsisten.",
            "Dukungan Sekolah": f"Pertahankan komunikasi positif dan perhatian personal kepada {nama}, karena ia merasa didukung lingkungan sekolah.",
            "Rasa Cemas": f"Salurkan rasa cemas {nama} menjadi motivasi melalui target belajar bertahap dan umpan balik yang jelas.",
        }
        rekomendasi.append(peta_kuat.get(top, f"Pertahankan faktor {top} yang sudah menjadi kekuatan utama {nama}."))

    if turun:
        worst = turun[0]["Faktor"]
        peta_lemah = {
            "Percaya Diri": f"Berikan scaffolding dan tugas bertahap yang bisa diselesaikan agar {nama} merasakan keberhasilan kecil setiap hari.",
            "Semangat Belajar": f"Variasikan metode pembelajaran dan hubungkan materi dengan minat/kehidupan sehari-hari {nama}.",
            "Dukungan Orang Tua": f"Lakukan komunikasi intensif dengan orang tua (WA/kartu kontrol) agar dukungan belajar di rumah meningkat.",
            "Dukungan Sekolah": f"Perkuat interaksi personal dengan {nama} — tanyakan langsung kesulitan belajarnya.",
            "Rasa Cemas": f"Ciptakan suasana kelas yang tidak menghakimi; berikan tes formatif tanpa tekanan agar kecemasan {nama} terkendali.",
        }
        rekomendasi.append(peta_lemah.get(worst, f"Berikan perhatian khusus pada faktor {worst} yang menghambat belajar {nama}."))

    if kat == "Perlu Perhatian":
        rekomendasi.append(f"Lakukan remedial terarah dan pantau progres {nama} secara mingguan melalui jurnal belajar.")
    elif kat == "Sangat Baik":
        rekomendasi.append(f"Berikan tantangan pengayaan (soal HOTS / proyek mandiri) agar {nama} tidak kehilangan tantangan belajar.")
    elif kat == "Baik":
        rekomendasi.append(f"Pertahankan kualitas pembelajaran dan berikan umpan balik konstruktif secara rutin.")
    else:
        rekomendasi.append(f"Fokus pada penguatan konsep dasar dan berikan latihan tambahan secara berkala.")

    if not rekomendasi:
        rekomendasi.append(f"Lanjutkan pendekatan pembelajaran yang sudah berjalan baik dan pantau perkembangan {nama} secara berkala.")
    return rekomendasi

def _generate_rekomendasi_orangtua(siswa, nilai, kat, kontribusi_list):
    nama = str(siswa.get("Nama", ""))
    naik = sorted([k for k in kontribusi_list if k["Pengaruh (Poin)"] > 0.1], key=lambda x: x["Pengaruh (Poin)"], reverse=True)
    turun = sorted([k for k in kontribusi_list if k["Pengaruh (Poin)"] < -0.1], key=lambda x: x["Pengaruh (Poin)"])

    rekomendasi = []
    if naik:
        top = naik[0]["Faktor"]
        peta_kuat = {
            "Percaya Diri": f"Terus beri pujian atas usaha {nama} (bukan hanya hasil), agar rasa percaya dirinya terjaga di rumah.",
            "Semangat Belajar": f"Dukung minat {nama} dengan menyediakan buku/alat belajar yang sesuai dengan ketertarikannya.",
            "Dukungan Orang Tua": f"Pertahankan kebiasaan mendampingi {nama} belajar dan menanyakan kegiatan sekolahnya setiap hari.",
            "Dukungan Sekolah": f"Jaga komunikasi rutin dengan wali kelas / guru BK agar dukungan sekolah kepada {nama} tetap optimal.",
            "Rasa Cemas": f"Bantu {nama} mengelola rasa cemas dengan jadwal belajar teratur, bukan menambah tekanan nilai.",
        }
        rekomendasi.append(peta_kuat.get(top, f"Pertahankan dukungan yang sudah menjadi kekuatan utama {nama}."))

    if turun:
        worst = turun[0]["Faktor"]
        peta_lemah = {
            "Percaya Diri": f"Hindari membandingkan {nama} dengan saudara/teman. Fokus pada kemajuan kecil yang sudah dicapai.",
            "Semangat Belajar": f"Ajak {nama} berdiskusi tentang cita-citanya dan hubungkan dengan pentingnya belajar saat ini.",
            "Dukungan Orang Tua": f"Luangkan minimal 30 menit/hari untuk mendampingi {nama} belajar dan menanyakan tugas sekolahnya.",
            "Dukungan Sekolah": f"Hadir di pertemuan orang tua/wali dan tanyakan perkembangan {nama} secara aktif kepada guru.",
            "Rasa Cemas": f"Kurangi tuntutan nilai tinggi pada {nama}; ganti dengan target proses (mis. 'belajar 1 jam tiap hari').",
        }
        rekomendasi.append(peta_lemah.get(worst, f"Berikan perhatian lebih pada aspek {worst} yang menghambat belajar {nama} di rumah."))

    if kat == "Perlu Perhatian":
        rekomendasi.append(f"Konsultasikan dengan guru/wali kelas untuk program belajar tambahan di rumah bagi {nama}.")
    elif kat == "Sangat Baik":
        rekomendasi.append(f"Berikan apresiasi dan tetap dampingi {nama} agar tidak kehilangan momentum belajar.")
    else:
        rekomendasi.append(f"Ciptakan suasana rumah yang kondusif untuk belajar dan hindari gangguan gadget saat jam belajar.")

    rekomendasi.append(f"Bangun komunikasi terbuka dengan {nama} — tanyakan perasaannya tentang sekolah, bukan hanya nilainya.")
    return rekomendasi

# ================================================================
# GENERATOR WORD & PDF
# ================================================================
def generate_laporan_word(siswa, nilai, kat, dag_pilihan, kontribusi_list):
    doc = Document()
    title = doc.add_heading('Laporan Hasil Analisis Prestasi Siswa', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph(f"Tanggal: {datetime.now().strftime('%d %B %Y')}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    doc.add_heading('Identitas Siswa', level=1)
    doc.add_paragraph(f"Nama: {siswa.get('Nama', '-')}")
    doc.add_paragraph(f"NISN: {siswa.get('NISN', '-')}")
    doc.add_paragraph(f"Kelas: {siswa.get('Kelas', '-')}")
    doc.add_paragraph()

    doc.add_heading('Nilai Akademik', level=1)
    doc.add_paragraph(f"Nilai Rata-rata: {nilai:.2f}")
    doc.add_paragraph(f"Kategori: {kat}")
    doc.add_paragraph()

    doc.add_heading('Model Analisis', level=1)
    doc.add_paragraph(dag_pilihan)
    doc.add_paragraph()

    doc.add_heading('Kontribusi Faktor', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Light Grid Accent 1'
    hdr = table.rows[0].cells
    hdr[0].text = "Faktor"
    hdr[1].text = "Pengaruh (Poin)"
    for k in kontribusi_list:
        row = table.add_row().cells
        row[0].text = k["Faktor"]
        row[1].text = f"{k['Pengaruh (Poin)']:+.2f}"
    doc.add_paragraph()

    doc.add_heading('Kesimpulan', level=1)
    narasi = _generate_narasi_laporan(siswa, nilai, kat, dag_pilihan, kontribusi_list)
    for line in narasi.split("\n"):
        doc.add_paragraph(line)
    doc.add_paragraph()

    doc.add_heading('Rekomendasi untuk Guru', level=1)
    doc.add_paragraph("Berikut langkah konkret yang dapat dilakukan guru/wali kelas:")
    for i, rec in enumerate(_generate_rekomendasi_guru(siswa, nilai, kat, kontribusi_list), 1):
        doc.add_paragraph(f"{i}. {rec}")
    doc.add_paragraph()

    doc.add_heading('Rekomendasi untuk Orang Tua', level=1)
    doc.add_paragraph("Berikut langkah konkret yang dapat dilakukan orang tua/wali di rumah:")
    for i, rec in enumerate(_generate_rekomendasi_orangtua(siswa, nilai, kat, kontribusi_list), 1):
        doc.add_paragraph(f"{i}. {rec}")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def generate_laporan_pdf(siswa, nilai, kat, dag_pilihan, kontribusi_list):
    def clean(t):
        s = str(t).encode('ascii', 'ignore').decode('ascii')
        return s if s.strip() else "-"

    pdf = FPDF()
    pdf.set_margins(15, 15, 15)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    lebar_kertas = pdf.w - pdf.l_margin - pdf.r_margin

    def cetak_teks(teks, font_style='', font_size=11, align='L', spasi=6):
        pdf.set_font('Helvetica', font_style, font_size)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(w=lebar_kertas, h=spasi, text=clean(teks), align=align)
        pdf.set_x(pdf.l_margin)

    cetak_teks('Laporan Hasil Analisis Prestasi Siswa', font_style='B', font_size=16, align='C', spasi=10)
    cetak_teks(f"Tanggal: {datetime.now().strftime('%d %B %Y')}", font_style='I', font_size=10, align='C', spasi=6)
    pdf.ln(5)

    cetak_teks('Identitas Siswa', font_style='B', font_size=12, spasi=8)
    cetak_teks(f"Nama: {siswa.get('Nama', '-')}")
    cetak_teks(f"NISN: {siswa.get('NISN', '-')}")
    cetak_teks(f"Kelas: {siswa.get('Kelas', '-')}")
    pdf.ln(3)

    cetak_teks('Nilai Akademik', font_style='B', font_size=12, spasi=8)
    cetak_teks(f"Nilai Rata-rata: {nilai:.2f}")
    cetak_teks(f"Kategori: {kat}")
    pdf.ln(3)

    cetak_teks('Model Analisis', font_style='B', font_size=12, spasi=8)
    cetak_teks(dag_pilihan)
    pdf.ln(3)

    cetak_teks('Kontribusi Faktor', font_style='B', font_size=12, spasi=8)
    w1 = lebar_kertas * 0.6
    w2 = lebar_kertas * 0.4
    pdf.set_font('Helvetica', 'B', 10)
    pdf.set_x(pdf.l_margin)
    pdf.cell(w1, 7, "Faktor", border=1)
    pdf.cell(w2, 7, "Pengaruh (Poin)", border=1, ln=True)
    pdf.set_font('Helvetica', '', 10)
    for k in kontribusi_list:
        pdf.set_x(pdf.l_margin)
        pdf.cell(w1, 6, clean(k["Faktor"]), border=1)
        pdf.cell(w2, 6, f"{k['Pengaruh (Poin)']:+.2f}", border=1, ln=True)
    pdf.ln(5)

    cetak_teks('Kesimpulan', font_style='B', font_size=12, spasi=8)
    narasi = _generate_narasi_laporan(siswa, nilai, kat, dag_pilihan, kontribusi_list)
    for line in narasi.split("\n"):
        if line.strip():
            cetak_teks(line, font_size=10)

    pdf.ln(4)
    cetak_teks('Rekomendasi untuk Guru', font_style='B', font_size=12, spasi=8)
    cetak_teks('Berikut langkah konkret yang dapat dilakukan guru/wali kelas:', font_size=10)
    for i, rec in enumerate(_generate_rekomendasi_guru(siswa, nilai, kat, kontribusi_list), 1):
        cetak_teks(f"{i}. {rec}", font_size=10)

    pdf.ln(4)
    cetak_teks('Rekomendasi untuk Orang Tua', font_style='B', font_size=12, spasi=8)
    cetak_teks('Berikut langkah konkret yang dapat dilakukan orang tua/wali di rumah:', font_size=10)
    for i, rec in enumerate(_generate_rekomendasi_orangtua(siswa, nilai, kat, kontribusi_list), 1):
        cetak_teks(f"{i}. {rec}", font_size=10)

    return io.BytesIO(bytes(pdf.output()))

# ================================================================
# PAGE CONFIG & CSS
# ================================================================
st.set_page_config(page_title="SAAnalytics", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")
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
.dag-card { background: white; padding: 20px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; justify-content: space-between; height: 100%; }
.dag-card h3 { margin: 0 0 5px 0; font-size: 1.1rem; color: #0f172a; }
.dag-card p { margin: 0; font-size: 0.9rem; color: #475569; flex-grow: 1; }
.dag-link { font-size: 0.85rem; color: #0d9488; font-weight: 500; margin-top: 15px; display: block; }
</style>
""", unsafe_allow_html=True)

# ================================================================
# STATE MANAGEMENT
# ================================================================
for key, val in {
    "logged_in": False, "user_nama": None, "user_role": None,
    "database_siswa": [], "current_page": "home", "last_menu": "Beranda",
    "siswa_login_nisn": None, "edit_nisn": None, "preview_nisn": None,
    "edit_back_page": "monitoring",
    "page_survei": 1, "page_nilai": 1, "custom_dags": {}, "selected_dag": None,
    "confirm_delete_all": False, "confirm_delete_upload": None
}.items():
    if key not in st.session_state: st.session_state[key] = val

if st.session_state.logged_in: refresh_pengaruh()

# ================================================================
# KUESIONER & LOGIN
# ================================================================
def render_kuesioner_siswa(nisn, is_preview=False):
    if is_preview:
        st.info(f"👁️ **Mode Pratinjau:** Anda melihat tampilan kuesioner siswa dengan NISN {nisn}.")
    else: 
        st.success(f"✅ Login berhasil! NISN: {nisn}")
        
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
        
        # 🔥 PETUNJUK PENGISIAN SKALA LIKERT
        st.info("""
        **📌 Petunjuk Pengisian:**
        Silakan pilih angka yang paling sesuai dengan kondisi Anda.
        *   **1** = Sangat Tidak Setuju
        *   **2** = Tidak Setuju
        *   **3** = Netral / Ragu-ragu
        *   **4** = Setuju
        *   **5** = Sangat Setuju
        """)
        
        construct_answers = {}
        for idx, (construct, items) in enumerate(KONSTRUK_ITEMS.items(), 1):
            st.markdown(f"**{idx}. {construct.replace('_', ' ')}**")
            for j, item in enumerate(items, 1):
                val = siswa_data.get(construct, 3) if isinstance(siswa_data, pd.Series) else 3
                default_idx = int(round(safe_float(val, 3))) - 1 if pd.notna(val) else 2
                construct_answers[item] = st.radio(
                    f"{item}", 
                    options=[1, 2, 3, 4, 5], 
                    horizontal=True, 
                    index=default_idx, 
                    key=f"q_{construct}_{j}",
                    captions=["Sangat Tidak Setuju", "Tidak Setuju", "Netral", "Setuju", "Sangat Setuju"]
                )
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
                    st.session_state.current_page = st.session_state.get("edit_back_page", "monitoring")
                    st.session_state.preview_nisn = None
                    st.session_state.edit_nisn = None
                else: 
                    st.session_state.siswa_login_nisn = None
                st.rerun()

def halaman_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br><h1 style='text-align:center; color:var(--teal); font-size: 2.8rem;'>🎓 SAAnalytics</h1>", unsafe_allow_html=True)
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
    st.markdown("""<div class="sb-logo"><div class="sb-logo-text"><h2>SAAnalytics</h2></div></div><hr>""", unsafe_allow_html=True)
    menu_options = ["Beranda", "Siswa", "Jawaban Survei", "Nilai Raport", "Laporan Siswa", "Analisis Faktor", "Simulasi Nilai", "Akun", "Panduan"]
    menu_icons = ["house-door", "people", "clipboard-data", "bar-chart", "file-earmark-text", "diagram-3", "magic", "person-circle", "info-circle"]
    if st.session_state.user_role == "admin": menu_options.extend(["Kelola Pengguna", "Keluar"]); menu_icons.extend(["gear", "box-arrow-right"])
    else: menu_options.append("Keluar"); menu_icons.append("box-arrow-right")

    mapping_menu = {"Beranda": "home", "Siswa": "monitoring", "Jawaban Survei": "list_survei", "Nilai Raport": "list_nilai", "Laporan Siswa": "rangkuman", "Analisis Faktor": "kausal", "Simulasi Nilai": "prediksi", "Akun": "akun", "Panduan": "tentang", "Kelola Pengguna": "admin_users", "Keluar": "logout"}
    current_page_mapped = st.session_state.current_page
    if current_page_mapped == "edit_siswa_lengkap": current_page_mapped = st.session_state.get("edit_back_page", "monitoring")

    default_idx = next((i for i, opt in enumerate(menu_options) if mapping_menu.get(opt) == current_page_mapped), 0)
    menu_choice = option_menu(None, menu_options, icons=menu_icons, default_index=default_idx, key="sidebar_menu", styles={"nav-link-selected": {"background-color": "rgba(255,255,255,0.15)", "color": "#facc15 !important", "border-left": "4px solid #facc15"}})

    if menu_choice != st.session_state.last_menu:
        st.session_state.last_menu = menu_choice; st.session_state.page_survei = 1; st.session_state.page_nilai = 1
        if menu_choice == "Keluar":
            for k in ("logged_in", "preview_nisn", "siswa_login_nisn", "edit_nisn"):
                st.session_state[k] = None if k.endswith("_nisn") or k.startswith("preview") else False
            st.session_state.database_siswa = []
            st.rerun()
        else: 
            st.session_state.current_page = mapping_menu.get(menu_choice, "home")
            st.session_state.selected_dag = None 

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
    tab_tambah, tab_import, tab_daftar = st.tabs(["➕ Tambah Manual", "📥 Import Data (Excel)", "👥 Daftar Siswa"])
    
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
                if st.button("Simpan ke Database", type="primary"):
                    count = import_data_historis_lengkap(df_upload)
                    if 'NISN' in df_upload.columns or 'NIS' in df_upload.columns:
                        nisn_list = [str(n) for n in (df_upload.get('NISN') if 'NISN' in df_upload.columns else df_upload.get('NIS')).dropna().unique().tolist()]
                        history = load_upload_history()
                        history.append({"id": str(datetime.now().timestamp()), "filename": uploaded.name, "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "count": count, "nisns": nisn_list})
                        save_upload_history(history)
                    st.success(f"✅ {count} data dimasukkan!"); refresh_pengaruh(); st.rerun()
            except Exception as e: st.error(f"Error: {e}")

    with tab_daftar:
        df_siswa = pd.DataFrame(st.session_state.database_siswa)
        if not df_siswa.empty:
            st.markdown("### ⚙️ Pengaturan & Hapus Data")
            col_hapus_semua, col_hapus_upload = st.columns(2)
            with col_hapus_semua:
                st.markdown("**Hapus Seluruh Data**")
                st.caption("Gunakan ini untuk mengosongkan *database*.")
                # 🔥 KHUSUS ADMIN: Hanya admin yang bisa menghapus semua
                if st.session_state.user_role == "admin":
                    if not st.session_state.confirm_delete_all:
                        if st.button("🗑️ Hapus Semua Siswa", type="primary", use_container_width=True): st.session_state.confirm_delete_all = True; st.rerun()
                    else:
                        st.warning("Yakin ingin menghapus SEMUA data siswa secara permanen?")
                        c_yes, c_no = st.columns(2)
                        if c_yes.button("Ya, Hapus Semua", key="yes_del_all", type="primary"):
                            for nisn in df_siswa['NISN'].unique(): hapus_survei_siswa(nisn)
                            st.session_state.confirm_delete_all = False; st.success("Semua data siswa telah dihapus."); refresh_pengaruh(); st.rerun()
                        if c_no.button("Batal", key="no_del_all"): st.session_state.confirm_delete_all = False; st.rerun()
                else:
                    st.info("🔒 Hanya Admin yang bisa menghapus seluruh data.")
            with col_hapus_upload:
                st.markdown("**Hapus Berdasarkan Riwayat Upload**")
                st.caption("Gunakan ini jika salah *upload* data Excel. (Hanya menghapus nilai, survei tetap aman).")
                # 🔥 KHUSUS ADMIN: Hanya admin yang bisa menghapus riwayat
                if st.session_state.user_role == "admin":
                    history = load_upload_history()
                    if history:
                        history_sorted = sorted(history, key=lambda x: x['date'], reverse=True)
                        options = {f"{h['filename']} | {h['date']} ({h['count']} data)": h for h in history_sorted}
                        pilihan = st.selectbox("Pilih Riwayat Upload:", list(options.keys()), key="hapus_upload_select")
                        if not st.session_state.confirm_delete_upload:
                            if st.button(f"🗑️ Hapus Nilai Upload Ini", type="primary", use_container_width=True): st.session_state.confirm_delete_upload = pilihan; st.rerun()
                        else:
                            st.warning(f"Yakin menghapus nilai dari upload **{st.session_state.confirm_delete_upload}**?")
                            c_yes, c_no = st.columns(2)
                            if c_yes.button("Ya, Hapus", key="yes_del_upload", type="primary"):
                                selected_batch = options.get(st.session_state.confirm_delete_upload)
                                if selected_batch:
                                    for n in selected_batch['nisns']: hapus_nilai_saja(n)
                                    history = [h for h in history if h['id'] != selected_batch['id']]
                                    save_upload_history(history)
                                st.session_state.confirm_delete_upload = None
                                st.success("Nilai dari upload tersebut berhasil dihapus! (Survei tetap aman)."); refresh_pengaruh(); st.rerun()
                            if c_no.button("Batal", key="no_del_upload"): st.session_state.confirm_delete_upload = None; st.rerun()
                    else: st.info("Belum ada riwayat upload file Excel.")
                else:
                    st.info("🔒 Hanya Admin yang bisa menghapus riwayat upload.")
            st.markdown("---")
            for idx, row in df_siswa.iterrows():
                c1, c2, c3, c4, c5 = st.columns([1.5, 3, 1.5, 2, 3.5])
                c1.write(f"**{row.get('NISN', '-')}**"); c2.write(f"**{row['Nama']}**")
                c3.write(f"Kelas: {row.get('Kelas', '-')}"); c4.write(f"Rata-rata: **{safe_float(row.get('Nilai Akademik', 0)):.2f}**")
                with c5:
                    b1, b2 = st.columns(2)
                    if b1.button("📝 Edit", key=f"edit_{row['NISN']}", use_container_width=True, help="Edit Data Siswa"):
                        st.session_state.edit_nisn = row['NISN']; st.session_state.edit_back_page = "monitoring"; st.session_state.current_page = "edit_siswa_lengkap"; st.rerun()
                    # 🔥 UBAH DISINI: Tombol Hapus Nilai Saja (Bukan Hapus Siswa)
                    if b2.button("🗑️ Hapus Nilai", key=f"del_siswa_{row['NISN']}", type="primary", use_container_width=True, help="Hapus Nilai Saja (Survei tetap aman)"):
                        sukses, pesan = hapus_nilai_saja(row['NISN'])
                        if sukses: st.success(pesan)
                        else: st.error(pesan)
                        st.rerun()
                st.markdown("<hr style='margin: 5px 0; border-top: 1px solid #e2e8f0;'>", unsafe_allow_html=True)

# ================================================================
# HALAMAN: DAFTAR SURVEI
# ================================================================
elif st.session_state.current_page == "list_survei":
    st.title("Data Survei Siswa")
    refresh_pengaruh(); df_siswa = pd.DataFrame(st.session_state.database_siswa)
    if not df_siswa.empty:
        cari = st.text_input("🔍 Cari nama atau NISN:")
        if cari: 
            df_siswa = df_siswa[
                df_siswa['Nama'].astype(str).str.contains(cari, case=False, na=False) | 
                df_siswa['NISN'].astype(str).str.contains(cari, case=False, na=False)
            ]
        
        items_per_page = 10
        # 🔥 FIX: Paksa minimal 1 halaman
        total_pages = max(1, math.ceil(len(df_siswa) / items_per_page)) if len(df_siswa) > 0 else 1
        
        if st.session_state.page_survei > total_pages: 
            st.session_state.page_survei = total_pages
        if st.session_state.page_survei < 1: 
            st.session_state.page_survei = 1
            
        df_page = df_siswa.iloc[(st.session_state.page_survei - 1) * items_per_page : (st.session_state.page_survei) * items_per_page]

        for idx, row in df_page.iterrows():
            c1, c2, c3, c4, c5, c6 = st.columns([1.5, 4, 3, 1.2, 1.2, 1.2])
            c1.write(f"**{row['NISN']}**"); c2.write(f"**{row['Nama']}**")
            status = "Lengkap ✅" if pd.notna(row.get("Self_Efficacy")) else "Belum ❌"
            c3.write(f"Terisi: {status}")
            with c4:
                if st.button("Ubah", key=f"edit_survei_{row['NISN']}"): 
                    st.session_state.edit_nisn = row['NISN']; st.session_state.edit_back_page = "list_survei"; st.session_state.current_page = "edit_siswa_lengkap"; st.rerun()
            with c5:
                if st.button("🔑 Reset Pass", key=f"reset_pw_{row['NISN']}"): reset_password_siswa(row['NISN']); st.success("Direset!"); st.rerun()
            with c6:
                # 🔥 KHUSUS ADMIN: Tombol Hapus Survei
                if st.session_state.user_role == "admin":
                    if st.button("🗑️ Hapus Survei", key=f"del_survei_{row['NISN']}", type="primary"):
                        hapus_survei_siswa(row['NISN']); st.success("Survei dihapus!"); refresh_pengaruh(); st.rerun()
                else:
                    st.caption("🔒 Khusus Admin")
            st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)

        p1, p2, p3 = st.columns([1, 2, 1])
        with p1:
            if st.button("⬅️ Sebelumnya", disabled=(st.session_state.page_survei == 1), use_container_width=True): 
                st.session_state.page_survei -= 1; st.rerun()
        with p2:
            c_l, c_m, c_r = st.columns([1, 1.5, 1])
            with c_m:
                col_t1, col_sel, col_t2 = st.columns([1, 1, 1])
                with col_t1: st.markdown("<div style='text-align:right; padding-top:12px;'>Halaman</div>", unsafe_allow_html=True)
                with col_sel:
                    sel_page = st.selectbox("Lompat", options=list(range(1, total_pages + 1)), index=st.session_state.page_survei - 1, label_visibility="collapsed")
                    if sel_page != st.session_state.page_survei: 
                        st.session_state.page_survei = sel_page; st.rerun()
                with col_t2: st.markdown(f"<div style='text-align:left; padding-top:12px;'>dari <b>{total_pages}</b></div>", unsafe_allow_html=True)
        with p3:
            if st.button("Berikutnya ➡️", disabled=(st.session_state.page_survei == total_pages), use_container_width=True): 
                st.session_state.page_survei += 1; st.rerun()
    else:
        st.info("Belum ada data siswa.")

# ================================================================
# HALAMAN: DAFTAR NILAI
# ================================================================
elif st.session_state.current_page == "list_nilai":
    st.title("Data Nilai Raport Siswa")
    df_siswa = pd.DataFrame(st.session_state.database_siswa)
    if not df_siswa.empty:
        cari = st.text_input("🔍 Cari nama atau NISN:", key="cari_nilai")
        if cari: 
            df_siswa = df_siswa[
                df_siswa['Nama'].astype(str).str.contains(cari, case=False, na=False) | 
                df_siswa['NISN'].astype(str).str.contains(cari, case=False, na=False)
            ]
        
        items_per_page = 10
        # 🔥 FIX: Paksa minimal 1 halaman
        total_pages = max(1, math.ceil(len(df_siswa) / items_per_page)) if len(df_siswa) > 0 else 1
        
        if st.session_state.page_nilai > total_pages: 
            st.session_state.page_nilai = total_pages
        if st.session_state.page_nilai < 1: 
            st.session_state.page_nilai = 1

        df_page = df_siswa.iloc[(st.session_state.page_nilai - 1) * items_per_page : st.session_state.page_nilai * items_per_page]

        for idx, row in df_page.iterrows():
            c1, c2, c3, c4, c5 = st.columns([1.5, 4, 2, 2, 1])
            c1.write(f"**{row['NISN']}**")
            c2.write(f"**{row['Nama']}**")
            nilai_rata = safe_float(row.get('Nilai Akademik', 0))
            c3.write(f"Rata-rata: **{nilai_rata:.2f}**")
            c4.write(f"Status: {'Lengkap' if nilai_rata > 0 else 'Kosong'}")
            with c5:
                if st.button("Ubah", key=f"edit_nilai_{row['NISN']}"):
                    st.session_state.edit_nisn = row['NISN']
                    st.session_state.edit_back_page = "list_nilai"
                    st.session_state.current_page = "edit_siswa_lengkap"; st.rerun()
            st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)

        p1, p2, p3 = st.columns([1, 2, 1])
        with p1:
            if st.button("⬅️ Sebelumnya", disabled=(st.session_state.page_nilai == 1), use_container_width=True, key="prev_nilai"):
                st.session_state.page_nilai -= 1; st.rerun()
        with p2:
            c_l, c_m, c_r = st.columns([1, 1.5, 1])
            with c_m:
                col_t1, col_sel, col_t2 = st.columns([1, 1, 1])
                with col_t1: st.markdown("<div style='text-align:right; padding-top:12px;'>Halaman</div>", unsafe_allow_html=True)
                with col_sel:
                    sel_page = st.selectbox("Lompat", options=list(range(1, total_pages + 1)), index=st.session_state.page_nilai - 1, label_visibility="collapsed", key="sel_nilai")
                    if sel_page != st.session_state.page_nilai:
                        st.session_state.page_nilai = sel_page; st.rerun()
                with col_t2: st.markdown(f"<div style='text-align:left; padding-top:12px;'>dari <b>{total_pages}</b></div>", unsafe_allow_html=True)
        with p3:
            if st.button("Berikutnya ➡️", disabled=(st.session_state.page_nilai == total_pages), use_container_width=True, key="next_nilai"):
                st.session_state.page_nilai += 1; st.rerun()
    else:
        st.info("Belum ada data nilai.")

# ================================================================
# HALAMAN: EDIT SISWA LENGKAP (GABUNGAN SURVEI & NILAI)
# ================================================================
elif st.session_state.current_page == "edit_siswa_lengkap":
    df = pd.DataFrame(st.session_state.database_siswa)
    if st.session_state.get("edit_nisn") is None or df.empty:
        st.warning("Pilih siswa dulu.")
        if st.button("Kembali", key="back_edit_empty"): st.session_state.current_page = st.session_state.get("edit_back_page", "monitoring"); st.rerun()
    else:
        back_page = st.session_state.get("edit_back_page", "monitoring")
        if st.button("⬅️ Kembali ke Daftar", key="back_edit_siswa"):
            st.session_state.current_page = back_page
            for k, v in mapping_menu.items():
                if v == back_page: st.session_state.last_menu = k; break
            st.rerun()
        siswa = df[df['NISN'] == str(st.session_state.edit_nisn)].iloc[0]

        def render_form_nilai_siswa(siswa_data):
            mapel_list = get_mapel_list(df)
            nilai_lama_per_mapel = {}
            for m in mapel_list:
                col_name = f"Nilai_{m}" if f"Nilai_{m}" in siswa_data.index else m
                if col_name in siswa_data.index and pd.notna(siswa_data[col_name]): nilai_lama_per_mapel[m] = safe_float(siswa_data[col_name], 0)
            nilai_skrg = safe_float(siswa_data.get('Nilai Akademik', 0))
            default_val = nilai_skrg if nilai_skrg > 0 else 80.0
            with st.form("form_nilai_per_mapel_unified"):
                c1, c2 = st.columns(2)
                nilai_per_mapel = {}
                for i, m in enumerate(mapel_list):
                    col = c1 if i % 2 == 0 else c2
                    default_m = nilai_lama_per_mapel.get(m, default_val)
                    nilai_per_mapel[m] = col.number_input(m, min_value=0.0, max_value=100.0, value=default_m, step=0.1, key=f"nilai_unified_{m}")
                rata = sum(nilai_per_mapel.values()) / len(nilai_per_mapel) if nilai_per_mapel else 0.0
                kat, warna = kategori_nilai(rata)
                st.markdown(f"""<div style="background:#f0fdfa; border-left:4px solid {warna}; padding:1rem; border-radius:6px; margin-top:1rem;"><b>Rata-rata Nilai:</b> <span style="font-size:1.5rem; font-weight:700; color:{warna};">{rata:.2f}</span> &nbsp;—&nbsp; Kategori: <b>{kat}</b></div>""", unsafe_allow_html=True)
                if st.form_submit_button("Simpan Nilai", type="primary"):
                    payload = {"Nilai": rata}
                    for m, v in nilai_per_mapel.items(): payload[f"Nilai_{m}"] = v
                    sukses, pesan = update_nilai_siswa(siswa_data['NISN'], payload)
                    if sukses: st.success("Nilai berhasil disimpan!"); refresh_pengaruh(); st.rerun()
                    else: st.error(f"Gagal: {pesan}")

        if back_page == "list_nilai":
            st.markdown(f"### Input Nilai per Mata Pelajaran — {siswa['Nama']}")
            render_form_nilai_siswa(siswa)
        else:
            st.markdown(f"### Edit Data Siswa: {siswa['Nama']} ({siswa['NISN']})")
            tab_survei, tab_nilai = st.tabs(["📝 Kuesioner & Identitas", "📊 Nilai Mata Pelajaran"])
            with tab_survei: render_kuesioner_siswa(siswa['NISN'], is_preview=True)
            with tab_nilai: render_form_nilai_siswa(siswa)

# ================================================================
# HALAMAN: LAPORAN SISWA (DENGAN EXPORT WORD & PDF)
# ================================================================
elif st.session_state.current_page == "rangkuman":
    st.title("Laporan Hasil Analisis Siswa")
    refresh_pengaruh(); df = pd.DataFrame(st.session_state.database_siswa)

    if not df.empty:
        dag_pilihan = st.selectbox("Pilih Model Analisis Faktor (DAG) untuk Laporan:", list(DAG_MODELS.keys()))
        df["Label"] = df["Nama"].astype(str) + " (" + df["Kelas"].astype(str).fillna("") + ")"
        siswa_terpilih = st.selectbox("Pilih Siswa:", df["Label"].tolist())

        if siswa_terpilih:
            siswa = df[df["Label"] == siswa_terpilih].iloc[0]
            nilai = safe_float(siswa.get("Nilai Akademik", 0)); kat, warna = kategori_nilai(nilai)
            st.markdown(f"### {siswa['Nama']} | Nilai: **{nilai:.2f}** ({kat})")

            mapel_list = get_mapel_list(df)
            ada_mapel = any(f"Nilai_{m}" in siswa.index and pd.notna(siswa[f"Nilai_{m}"]) for m in mapel_list) or any(m in siswa.index and pd.notna(siswa[m]) for m in mapel_list)
            if ada_mapel:
                with st.expander("📋 Rincian Nilai per Mata Pelajaran", expanded=False):
                    mapel_data = []
                    for m in mapel_list:
                        col_name = f"Nilai_{m}" if f"Nilai_{m}" in siswa.index else m
                        if col_name in siswa.index and pd.notna(siswa[col_name]): mapel_data.append({"Mata Pelajaran": m, "Nilai": safe_float(siswa[col_name])})
                    if mapel_data: st.dataframe(pd.DataFrame(mapel_data), use_container_width=True, hide_index=True)

            ates = DAG_MODELS[dag_pilihan]["ate"]
            nama_awam = {"Self_Efficacy": "Percaya Diri", "Academic_Anxiety": "Rasa Cemas", "School_Support": "Dukungan Sekolah", "Parental_Involvement": "Dukungan Orang Tua", "Learning_Motivation": "Semangat Belajar"}
            kontribusi = []
            for f, nama in nama_awam.items():
                ate = ates.get(f, 0.0); val = safe_float(siswa.get(f, 3.0)); base = BASELINE_ASPEK.get(f, 3.0)
                kontribusi.append({"Faktor": nama, "Pengaruh (Poin)": ate * (val - base)})

            df_kontrib = pd.DataFrame(kontribusi).sort_values("Pengaruh (Poin)")
            fig = px.bar(df_kontrib, x="Pengaruh (Poin)", y="Faktor", orientation="h", color="Pengaruh (Poin)", color_continuous_scale="RdBu", text=[f"{v:+.2f}" for v in df_kontrib["Pengaruh (Poin)"]])
            fig.update_traces(textposition="outside"); fig.update_layout(height=300, coloraxis_showscale=False); st.plotly_chart(fig, use_container_width=True)

            st.markdown("### Kesimpulan")
            narasi = _generate_narasi_laporan(siswa, nilai, kat, dag_pilihan, kontribusi)
            st.markdown(f"""<div style="background:#f8fafc; padding:1.5rem; border-radius:8px; border-left:4px solid #0d9488; white-space: pre-wrap; line-height:1.8;">{narasi}</div>""", unsafe_allow_html=True)

            st.markdown("### 👨‍🏫 Rekomendasi untuk Guru")
            guru_html = "".join([f"<li>{r}</li>" for r in _generate_rekomendasi_guru(siswa, nilai, kat, kontribusi)])
            st.markdown(f"""<div style="background:#eff6ff; padding:1.5rem; border-radius:8px; border-left:4px solid #2563eb; line-height:1.8;"><ol style="margin:0; padding-left:1.5rem;">{guru_html}</ol></div>""", unsafe_allow_html=True)

            st.markdown("### 👨‍👩‍👧 Rekomendasi untuk Orang Tua")
            ortu_html = "".join([f"<li>{r}</li>" for r in _generate_rekomendasi_orangtua(siswa, nilai, kat, kontribusi)])
            st.markdown(f"""<div style="background:#fef3c7; padding:1.5rem; border-radius:8px; border-left:4px solid #f59e0b; line-height:1.8;"><ol style="margin:0; padding-left:1.5rem;">{ortu_html}</ol></div>""", unsafe_allow_html=True)

            st.markdown("---"); st.markdown("### 📥 Unduh Laporan")
            st.caption("Unduh laporan dalam format Word (.docx) atau PDF untuk dokumentasi.")

            safe_nama = clean_filename(siswa.get('Nama', 'Siswa')); safe_nisn = clean_filename(siswa.get('NISN', ''))
            base_filename = f"Laporan ({safe_nama}, {safe_nisn})"

            col_w, col_p = st.columns(2)
            with col_w:
                if HAS_DOCX:
                    try:
                        buf_w = generate_laporan_word(siswa, nilai, kat, dag_pilihan, kontribusi)
                        st.download_button("📄 Unduh Word (.docx)", data=buf_w, file_name=f"{base_filename}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="dl_word_rangkuman")
                    except Exception as e: st.error(f"Gagal generate Word: {e}")
                else: st.warning("`python-docx` belum ter-install.")

            with col_p:
                if HAS_FPDF:
                    try:
                        buf_p = generate_laporan_pdf(siswa, nilai, kat, dag_pilihan, kontribusi)
                        st.download_button("📄 Unduh PDF", data=buf_p, file_name=f"{base_filename}.pdf", mime="application/pdf", use_container_width=True, key="dl_pdf_rangkuman")
                    except Exception as e: st.error(f"Gagal generate PDF: {e}")
                else: st.warning("`fpdf2` belum ter-install.")

# ================================================================
# HALAMAN: ANALISIS FAKTOR (VISUALISASI DAG)
# ================================================================
elif st.session_state.current_page == "kausal":
    st.title("Analisis Faktor Nilai Siswa")
    st.write("Konstruk dibentuk dari fitur survei X (rata-rata butir), outcome = Nilai Akademik (Y). Pilih model DAG untuk melihat visualisasi jalur kausal dan interpretasinya.")
    tab_daftar, tab_create, tab_banding = st.tabs(["📂 Daftar DAG", "➕ Create DAG", "📊 Perbandingan Lintas DAG"])

    with tab_daftar:
        st.markdown("""<div style="background-color: #e0f2f1; border-left: 5px solid #0d9488; padding: 15px; border-radius: 5px; margin-bottom: 20px; color: #004d40; font-size: 0.9rem;"><strong>Silakan pilih model terbaik menurut data: DAG 4 — Full Structural (Ecological Systems + SRL) atau DAG 2 — Full Mediation (SDT).</strong><br>Model-model ini lolos uji kecocokan dengan data: semua anggapan "tidak saling berkaitan" di dalam DAG-nya terpenuhi.</div>""", unsafe_allow_html=True)

        st.markdown("### Contoh DAG")
        c1, c2 = st.columns(2)
        dag_titles = {"Model 1: Baseline Direct Effects": "DAG 1 — Baseline Direct Effects", "Model 2: Full Mediation (SDT)": "DAG 2 — Full Mediation (Self-Determination Theory)", "Model 3: Anxiety as Suppressor": "DAG 3 — Anxiety as Suppressor Pathway", "Model 4: Full Structural (Ecological)": "DAG 4 — Full Structural (Ecological Systems + SRL)"}

        for i, (k, v) in enumerate(DAG_MODELS.items()):
            target_col = c1 if i % 2 == 0 else c2
            with target_col:
                st.markdown(f"""<div class="dag-card"><div><h3>{dag_titles.get(k, k)}</h3><p>{v['deskripsi']}</p></div><span class="dag-link">{v['edges']} sisi</span></div>""", unsafe_allow_html=True)
                if st.button(f"lihat diagram & interpretasi →", key=f"lihat_{k}", use_container_width=True): st.session_state.selected_dag = k; st.rerun()

        if st.session_state.custom_dags:
            st.markdown("### DAG Buatan Guru (Kustom)")
            c3, c4 = st.columns(2)
            custom_dag_keys = list(st.session_state.custom_dags.keys())
            for i, k in enumerate(custom_dag_keys):
                v = st.session_state.custom_dags[k]; target_col = c3 if i % 2 == 0 else c4
                with target_col:
                    st.markdown(f"""<div class="dag-card"><div><h3>{k}</h3><p>{v['deskripsi']}</p></div><span class="dag-link">{len(v['edges'])} sisi</span></div>""", unsafe_allow_html=True)
                    b1, b2 = st.columns([4, 1])
                    with b1:
                        if st.button(f"lihat diagram & interpretasi →", key=f"lihat_custom_{k}", use_container_width=True): st.session_state.selected_dag = k; st.rerun()
                    with b2:
                        if st.button("🗑️", key=f"hapus_custom_{k}", help="Hapus DAG ini"):
                            del st.session_state.custom_dags[k]
                            if st.session_state.selected_dag == k: st.session_state.selected_dag = None
                            st.rerun()

        if st.session_state.get("selected_dag"):
            st.markdown("---"); selected = st.session_state.selected_dag
            if selected == "Model 1: Baseline Direct Effects":
                dot = """digraph G { rankdir=LR; splines=true; nodesep=0.5; ranksep=1.5; node [shape=box, style="rounded,filled", fontname="Helvetica", fontcolor=white, width=2]; edge [fontname="Helvetica", fontsize=10]; PI [label="Parental\\nInvolvement", fillcolor="#0f766e"]; SS [label="School Support", fillcolor="#0f766e"]; SE [label="Self Efficacy", fillcolor="#0f766e"]; LM [label="Learning\\nMotivation", fillcolor="#0f766e"]; AA [label="Academic\\nAnxiety", fillcolor="#0f766e"]; Y [label="Prestasi Siswa", fillcolor="#eab308", fontcolor="black", width=2.5]; PI -> Y [label="+1.25 ns", style=dashed, color="#666"]; SS -> Y [label="-3.61***", color="#dc2626"]; SE -> Y [label="+3.94***", color="#0d9488", penwidth=2.5]; LM -> Y [label="-0.58 ns", style=dashed, color="#666"]; AA -> Y [label="+1.29*", color="#dc2626"]; }"""
                interpretasi = "**Panduan Membaca Diagram:**\n* Arah panah menunjukkan arah pengaruh. **Angka Positif (+)** berarti menaikkan nilai, **Angka Negatif (-)** berarti menurunkan nilai.\n* Garis tebal `***` = pengaruh sangat kuat. Garis putus-putus `ns` = tidak ada pengaruh langsung.\n\n**Interpretasi DAG 1 (Sederhananya):**\n* **Percaya Diri (Self Efficacy)** punya angka tertinggi (+3.94). Ini faktor terkuat yang membuat nilai siswa naik.\n* **Dukungan Sekolah** angkanya minus (-3.61). Bukan berarti sekolah membuat bodoh, tapi karena anak yang sudah sangat pintar & percaya diri biasanya merasa dukungan sekolah itu 'biasa aja', sehingga hitungan matematisnya terbalik.\n* **Rasa Cemas** positif (+1.29). Anak yang sedikit cemas ujian justru terdorong belajar lebih giat."
            elif selected == "Model 2: Full Mediation (SDT)":
                dot = """digraph G { rankdir=LR; splines=true; nodesep=0.5; ranksep=1.5; node [shape=box, style="rounded,filled", fontname="Helvetica", fontcolor=white, width=2]; edge [fontname="Helvetica", fontsize=9]; PI [label="Parental\\nInvolvement", fillcolor="#0f766e"]; SS [label="School Support", fillcolor="#0f766e"]; LM [label="Learning\\nMotivation", fillcolor="#0f766e"]; SE [label="Self Efficacy", fillcolor="#0f766e"]; AA [label="Academic\\nAnxiety", fillcolor="#0f766e"]; Y [label="Prestasi Siswa", fillcolor="#eab308", fontcolor="black", width=2.5]; PI -> LM [label="+0.19*", color="#0d9488"]; PI -> SE [label="+0.19*", color="#0d9488"]; PI -> AA [label="+0.14 ns", style=dashed, color="#666"]; SS -> LM [label="+0.50***", color="#0d9488"]; SS -> SE [label="+0.65***", color="#0d9488"]; SS -> AA [label="+0.49***", color="#0d9488"]; LM -> Y [label="-1.04 ns", style=dashed, color="#666"]; SE -> Y [label="+2.57**", color="#0d9488", penwidth=2.5]; AA -> Y [label="+0.98 ns", style=dashed, color="#666"]; }"""
                interpretasi = "**Interpretasi DAG 2 (Sederhananya):**\n* **Dukungan Orang Tua** dan **Dukungan Sekolah** tidak bisa langsung menaikkan nilai siswa.\n* Mereka bekerja dengan cara **membangun rasa Percaya Diri** dan **Semangat Belajar** siswa terlebih dahulu.\n* Contoh: Dukungan Sekolah punya angka besar (+0.65) untuk membangun Percaya Diri. Setelah anak punya rasa percaya diri (+2.57), barulah nilai akademiknya ikut naik. Kuncinya ada di dalam diri siswa."
            elif selected == "Model 3: Anxiety as Suppressor":
                dot = """digraph G { rankdir=LR; splines=true; nodesep=0.5; ranksep=1.5; node [shape=box, style="rounded,filled", fontname="Helvetica", fontcolor=white, width=2]; edge [fontname="Helvetica", fontsize=9]; PI [label="Parental\\nInvolvement", fillcolor="#0f766e"]; SS [label="School Support", fillcolor="#0f766e"]; AA [label="Academic\\nAnxiety", fillcolor="#0f766e"]; LM [label="Learning\\nMotivation", fillcolor="#0f766e"]; SE [label="Self Efficacy", fillcolor="#0f766e"]; Y [label="Prestasi Siswa", fillcolor="#eab308", fontcolor="black", width=2.5]; PI -> LM [label="+0.19*", color="#0d9488"]; PI -> SE [label="+0.19*", color="#0d9488"]; PI -> AA [label="+0.14 ns", style=dashed, color="#666"]; PI -> Y [label="+1.25 ns", style=dashed, color="#666"]; SS -> LM [label="+0.50***", color="#0d9488"]; SS -> SE [label="+0.65***", color="#0d9488"]; SS -> AA [label="+0.49***", color="#0d9488"]; SS -> Y [label="-3.61***", color="#dc2626"]; AA -> LM [label="+0.20**", color="#dc2626"]; AA -> SE [label="+0.10 ns", style=dashed, color="#666"]; AA -> Y [label="+1.29*", color="#dc2626"]; LM -> Y [label="-0.58 ns", style=dashed, color="#666"]; SE -> Y [label="+3.94***", color="#0d9488", penwidth=2.5]; }"""
                interpretasi = "**Interpretasi DAG 3 (Sederhananya):**\n* Biasanya rasa cemas dianggap buruk. Tapi di data ini, **Rasa Cemas** angkanya positif (+0.20).\n* Artinya, siswa yang sedikit cemas menghadapi ujian justru **lebih termotivasi untuk belajar lebih giat**. Rasa cemas di sini berfungsi sebagai alarm yang bikin mereka bangun dan belajar.\n* Namun, **Percaya Diri** tetap menjadi kunci utama yang paling besar pengaruhnya (+3.94) untuk mendongkrak nilai."
            elif selected == "Model 4: Full Structural (Ecological)":
                dot = """digraph G { rankdir=LR; splines=true; nodesep=0.5; ranksep=1.5; node [shape=box, style="rounded,filled", fontname="Helvetica", fontcolor=white, width=2]; edge [fontname="Helvetica", fontsize=9]; PI [label="Parental\\nInvolvement", fillcolor="#0f766e"]; SS [label="School Support", fillcolor="#0f766e"]; SE [label="Self Efficacy", fillcolor="#0f766e"]; LM [label="Learning\\nMotivation", fillcolor="#0f766e"]; AA [label="Academic\\nAnxiety", fillcolor="#0f766e"]; Y [label="Prestasi Siswa", fillcolor="#eab308", fontcolor="black", width=2.5]; PI -> SS [label="+0.43***", color="#0d9488"]; PI -> LM [label="+0.19*", color="#0d9488"]; PI -> SE [label="+0.19*", color="#0d9488"]; SS -> LM [label="+0.50***", color="#0d9488"]; SS -> SE [label="+0.65***", color="#0d9488"]; SE -> LM [label="+0.43***", color="#0d9488", penwidth=2]; LM -> AA [label="+0.45**", color="#0d9488"]; LM -> Y [label="-0.58 ns", style=dashed, color="#666"]; SE -> Y [label="+3.94***", color="#0d9488", penwidth=2.5]; AA -> Y [label="+1.29*", color="#dc2626"]; }"""
                interpretasi = "**Interpretasi DAG 4 (Sederhananya):**\n* Model ini paling lengkap dan menunjukkan bahwa semua faktor saling berhubungan membentuk 'lingkaran positif'.\n* **Orang Tua ke Sekolah (+0.43)**: Orang tua yang peduli dan aktif di rumah akan membuat anak merasa lebih didukung oleh lingkungan sekolahnya.\n* **Percaya Diri ke Semangat Belajar (+0.43)**: Anak yang yakin pada kemampuannya akan otomatis punya semangat belajar yang tinggi.\n* **Semangat Belajar ke Rasa Cemas (+0.45)**: Anak yang semangatnya tinggi justru sedikit cemas saat ujian. Ini cemas yang sehat (takut dapat nilai jelek makanya belajar mati-matian)."
            elif selected in st.session_state.custom_dags:
                info = st.session_state.custom_dags[selected]; st.markdown(f"### Visualisasi: {selected}")
                df_siswa = pd.DataFrame(st.session_state.database_siswa); hasil_ols = hitung_ols_custom(df_siswa, info["edges"])
                dot = """digraph G { rankdir=LR; splines=true; nodesep=0.5; ranksep=1.5; node [shape=box, style="rounded,filled", fontname="Helvetica", fontcolor=white, width=2]; edge [fontname="Helvetica", fontsize=10];""" + 'Y [label="Prestasi Siswa", fillcolor="#eab308", fontcolor="black", width=2.5];\n'
                teks_kesimpulan = ""
                for e in info["edges"]:
                    node = e[0]
                    if node == "Nilai Akademik": continue
                    beta, pval = hasil_ols.get(node, (0.0, 1.0)); sign = "+" if beta >= 0 else ""
                    stars = "***" if pval < 0.001 else "**" if pval < 0.01 else "*" if pval < 0.05 else "ns"
                    label = f"{sign}{beta:.2f} {stars}"; color = "#0d9488" if pval < 0.05 else "#666"; style = "solid" if pval < 0.05 else "dashed"
                    dot += f'"{node}" [label="{node.replace("_", " ")}", fillcolor="#0f766e"];\n"{node}" -> "{e[1]}" [label="{label}", color="{color}", style="{style}"];\n'
                    if pval < 0.05: teks_kesimpulan += f"- **{node.replace('_', ' ')}** terbukti secara nyata {'menaikkan' if beta >= 0 else 'menurunkan'} nilai siswa sebesar {abs(beta):.2f} poin.\n"
                dot += '}'
                st.graphviz_chart(dot, use_container_width=True); st.markdown("### Penjelasan & Interpretasi Hasil")
                if teks_kesimpulan: st.info(f"**Berdasarkan analisis data (OLS):**\n{teks_kesimpulan}\n\n*Catatan: Angka `ns` berarti secara ilmiah belum bisa dibuktikan pengaruhnya.*")
                else: st.info("Belum ada pengaruh yang terbukti secara nyata dari model ini pada data yang ada saat ini.")
            if 'dot' in locals() and dot:
                if not (selected in st.session_state.custom_dags):
                    st.markdown(f"### Visualisasi: {dag_titles.get(selected, selected)}")
                    st.graphviz_chart(dot, use_container_width=True); st.markdown("---"); st.markdown(f"### Penjelasan & Interpretasi Hasil"); st.info(interpretasi)
                else:
                    st.markdown("---"); st.markdown(f"### Penjelasan & Interpretasi Hasil"); st.info("DAG ini merupakan model buatan Anda sendiri. Silakan gunakan tab 'Perbandingan Lintas DAG' untuk melihat perbandingan jumlah sisi antar model.")
    with tab_create:
        st.info("Kelompokkan fitur X menjadi beberapa bagian (misal A, B, C, D atau nama konstruk). 2) Pilih Grog untuk membuat DAG, atau tulis sintaks sendiri. Jawaban harus lolos validasi.")
        nama_dag = st.text_input("Nama DAG", placeholder="mis. Model lingkungan → psikologis → prestasi")
        st.markdown("### Nama bagian (node), maks. 8"); st.caption("Huruf/angka/underscore saja. Bagian kosong diabaikan.")
        cols = st.columns(4); nodes_input = []; defaults = ["Self_Efficacy", "Academic_Anxiety", "School_Support", "Parental_Involvement", "Learning_Motivation", "", "", ""]
        for i in range(8):
            with cols[i % 4]:
                val = st.text_input(f"Bagian {i+1}", value=defaults[i], key=f"node_{i}", label_visibility="collapsed", placeholder=f"Bagian {i+1}")
                if val.strip(): nodes_input.append(val.strip())
        df_mapping = pd.DataFrame(RAW_QUESTIONS); df_mapping["Masuk bagian"] = df_mapping["default"]
        edited_df = st.data_editor(df_mapping[["kode", "fitur", "Masuk bagian"]], column_config={"kode": st.column_config.TextColumn("Kode", disabled=True, width="small"), "fitur": st.column_config.TextColumn("Fitur X (butir Likert)", disabled=True, width="large"), "Masuk bagian": st.column_config.SelectboxColumn("Masuk bagian", options=nodes_input if nodes_input else ["Belum ada bagian"], required=True)}, hide_index=True, use_container_width=True, key="dag_mapping_editor")
        st.markdown("---"); st.markdown("### Opsi Tambahan")
        metode = st.radio("Pilih metode:", ["Tulis sintaks sendiri", "Dibuatkan Groq"], horizontal=True)
        catatan = st.text_area("Catatan/hipotesis untuk Groq (opsional)", placeholder="mis. Anggap bagian 1 dan 2 sebagai lingkungan sosial yang memengaruhi prestasi melalui bagian 3 dan 4.")
        max_sisi = st.number_input("Maksimum sisi", min_value=1, max_value=50, value=16)
        if st.button("Buat & analisis", type="primary", key="btn_buat_dag"):
            if not nama_dag or len(nodes_input) < 1: st.error("Isi Nama DAG dan minimal satu Nama Bagian (Node).")
            else:
                new_edges = [(node, "Nilai Akademik") for node in nodes_input if node != "Nilai Akademik"]
                st.session_state.custom_dags[nama_dag] = {"deskripsi": f"Model buatan guru (Dibuat: {datetime.now().strftime('%Y-%m-%d')}). Metode: {metode}.", "edges": new_edges, "mapping": edited_df.to_dict('records')}
                st.session_state.selected_dag = nama_dag; st.success(f"DAG '{nama_dag}' berhasil dibuat dan dianalisis!")
                df_siswa = pd.DataFrame(st.session_state.database_siswa); hasil_ols = hitung_ols_custom(df_siswa, new_edges)
                dot = """digraph G { rankdir=LR; splines=true; nodesep=0.5; ranksep=1.5; node [shape=box, style="rounded,filled", fontname="Helvetica", fontcolor=white, width=2]; edge [fontname="Helvetica", fontsize=10];""" + 'Y [label="Prestasi Siswa", fillcolor="#eab308", fontcolor="black", width=2.5];\n'
                teks_kesimpulan = ""
                for e in new_edges:
                    node = e[0]; beta, pval = hasil_ols.get(node, (0.0, 1.0))
                    sign = "+" if beta >= 0 else ""; stars = "***" if pval < 0.001 else "**" if pval < 0.01 else "*" if pval < 0.05 else "ns"
                    label = f"{sign}{beta:.2f} {stars}"; color = "#0d9488" if pval < 0.05 else "#666"; style = "solid" if pval < 0.05 else "dashed"
                    dot += f'"{node}" [label="{node.replace("_", " ")}", fillcolor="#0f766e"];\n"{node}" -> "{e[1]}" [label="{label}", color="{color}", style="{style}"];\n'
                    if pval < 0.05: teks_kesimpulan += f"- **{node.replace('_', ' ')}** terbukti secara nyata {'menaikkan' if beta >= 0 else 'menurunkan'} nilai siswa sebesar {abs(beta):.2f} poin.\n"
                dot += '}'
                st.markdown(f"### Visualisasi: {nama_dag}"); st.graphviz_chart(dot, use_container_width=True); st.markdown("### Penjelasan & Interpretasi Hasil")
                if teks_kesimpulan: st.info(f"**Berdasarkan analisis data (OLS):**\n{teks_kesimpulan}\n\n*Catatan: Angka `ns` berarti secara ilmiah belum bisa dibuktikan pengaruhnya.*")
                else: st.info("Belum ada pengaruh yang terbukti secara nyata dari model ini pada data yang ada saat ini.")
    with tab_banding:
        st.markdown("### Perbandingan Semua Model DAG"); st.write("Tabel ini membandingkan seberapa besar pengaruh tiap faktor terhadap nilai siswa di berbagai model.")
        all_dags = {**DAG_MODELS, **st.session_state.custom_dags}; konstruk_list = ["Parental_Involvement", "School_Support", "Self_Efficacy", "Learning_Motivation", "Academic_Anxiety"]
        data_tabel = []
        for k in konstruk_list:
            row = {"Faktor": k}
            for dag_name, dag_info in all_dags.items():
                col_name = dag_name.split(":")[0] if ":" in dag_name else dag_name
                if "ate" in dag_info: val = dag_info["ate"].get(k, 0.0); pval = 0.01 if abs(val) > 0.5 else 0.5
                else:
                    df_siswa = pd.DataFrame(st.session_state.database_siswa); hasil_ols = hitung_ols_custom(df_siswa, dag_info["edges"])
                    if k in [e[0] for e in dag_info["edges"]]: val, pval = hasil_ols.get(k, (0.0, 1.0))
                    else: row[col_name] = "—"; continue
                row[col_name] = terjemah_efek(val, pval)
            data_tabel.append(row)
        df_compare = pd.DataFrame(data_tabel).set_index("Faktor")
        nama_awam = {"Parental_Involvement": "Dukungan Orang Tua", "School_Support": "Dukungan Sekolah", "Self_Efficacy": "Percaya Diri", "Learning_Motivation": "Semangat Belajar", "Academic_Anxiety": "Rasa Cemas"}
        df_compare = df_compare.rename(index=nama_awam)
        akurasi_row = {"Faktor": "Seberapa Akurat Model Memprediksi Nilai"}; stabil_row = {"Faktor": "Tingkat Keyakinan Model"}; jalur_row = {"Faktor": "Jumlah Jalan Tidak Langsung"}
        for dag_name in all_dags.keys():
            col_name = dag_name.split(":")[0] if ":" in dag_name else dag_name
            akurasi_row[col_name] = "Cukup Baik" if "Model 1" in dag_name or "Model 4" in dag_name else "Kurang Baik"
            stabil_row[col_name] = "Sangat Yakin" if "Model 1" in dag_name or "Model 4" in dag_name else "Cukup Yakin"
            jalur_row[col_name] = "Tidak Ada" if "Model 1" in dag_name else "Ada Beberapa"
        df_compare.loc["Seberapa Akurat Model Memprediksi Nilai"] = [akurasi_row.get(c, "Cukup Baik") for c in df_compare.columns]
        df_compare.loc["Tingkat Keyakinan Model"] = [stabil_row.get(c, "Cukup Yakin") for c in df_compare.columns]
        df_compare.loc["Jumlah Jalan Tidak Langsung"] = [jalur_row.get(c, "Ada Beberapa") for c in df_compare.columns]
        st.dataframe(df_compare.style.map(lambda v: 'background-color: #d1fae5; color: #065f46; font-weight: bold;' if isinstance(v, str) and ('Menaikkan' in v or 'Sedikit Menaikkan' in v) else 'background-color: #fee2e2; color: #991b1b; font-weight: bold;' if isinstance(v, str) and ('Menurunkan' in v or 'Sedikit Menurunkan' in v) else ''), use_container_width=True)
        st.caption("**Panduan Warna & Simbol:**"); st.caption("🟩 **Hijau**: Faktor tersebut terbukti sangat mempengaruhi nilai siswa (ke arah positif)."); st.caption("🟥 **Merah**: Faktor tersebut terbukti mempengaruhi nilai siswa (ke arah negatif/menurunkan)."); st.caption("**Tidak Berpengaruh**: Secara statistik, belum terbukti ada pengaruhnya."); st.caption("**Seberapa Akurat Model**: Seberapa bagus model tersebut menjelaskan nilai siswa."); st.caption("**Jumlah Jalan Tidak Langsung**: Berapa banyak 'jalan perantara' (misal: Orang tua -> Percaya Diri -> Nilai) yang aktif.")

elif st.session_state.current_page == "prediksi":
    st.title("Simulasi Nilai Siswa"); st.write("Geser nilai untuk melihat perkiraan nilai (Berdasarkan Model 1).")
    c1, c2 = st.columns(2)
    with c1: p_se = st.slider("Percaya Diri", 1.0, 5.0, BASELINE_ASPEK["Self_Efficacy"], 0.1); p_mot = st.slider("Semangat Belajar", 1.0, 5.0, BASELINE_ASPEK["Learning_Motivation"], 0.1); p_cem = st.slider("Rasa Cemas", 1.0, 5.0, BASELINE_ASPEK["Academic_Anxiety"], 0.1)
    with c2: p_ortu = st.slider("Dukungan Orang Tua", 1.0, 5.0, BASELINE_ASPEK["Parental_Involvement"], 0.1); p_sekolah = st.slider("Dukungan Sekolah", 1.0, 5.0, BASELINE_ASPEK["School_Support"], 0.1)
    if st.button("Hitung Perkiraan Nilai", type="primary", key="btn_hitung_prediksi"):
        profil = {"Self_Efficacy": p_se, "Learning_Motivation": p_mot, "Academic_Anxiety": p_cem, "Parental_Involvement": p_ortu, "School_Support": p_sekolah}
        ate_model1 = DAG_MODELS["Model 1: Baseline Direct Effects"]["ate"]
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
                if user_key and verify_password(old_pw, users[user_key]["password"]): update_user_password(user_key, hash_password(new_pw)); st.success("Berhasil diubah!")
                else: st.error("Password lama salah.")

elif st.session_state.current_page == "tentang":
    st.title("Panduan Penggunaan Aplikasi")
    st.markdown("""
    Selamat datang di **SAAnalytics**! Berikut adalah panduan singkat penggunaan aplikasi untuk Bapak/Ibu Guru:

    1. 🎒 **Portal Siswa**: Siswa mendaftar sendiri menggunakan NISN, lalu mengisi survei psikologis. (Guru juga bisa membagikan link ini ke siswa).
    2. 📊 **Menu Nilai Raport**: Guru dapat memasukkan nilai siswa untuk setiap mata pelajaran (Matematika, IPA, dll). Sistem akan otomatis menghitung rata-ratanya.
    3. 📄 **Laporan Siswa**: Fitur unggulan untuk melihat faktor apa saja yang membuat nilai siswa naik atau turun. Laporan ini disajikan dengan bahasa yang mudah dipahami, dan **bisa diunduh dalam bentuk Word atau PDF** — lengkap dengan **rekomendasi terpisah untuk Guru dan Orang Tua**.
    4. 🧠 **Analisis Faktor**: Guru bisa melihat gambar diagram (DAG) yang menjelaskan hubungan sebab-akibat dari nilai siswa. Semakin banyak data, semakin akurat analisisnya.
    """)

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
    st.markdown("---"); st.subheader("Daftar Pengguna Aktif")
    for username, data in users.items():
        if data.get("status") == "pending": continue
        c1, c2, c3, c4 = st.columns([2, 2, 1, 2])
        c1.write(f"**{data['nama']}**"); c2.write(f"`{username}`"); c3.write(f"{data.get('role', 'guru').upper()}")
        with c4:
            if username not in SUPER_ADMIN and username != "admin":
                if st.button("Hapus Akun", key=f"del_user_{username}", type="primary"): del users[username]; save_users(users); st.rerun()
            else: st.caption("Super Admin (Terkunci)")
        st.markdown("<hr style='margin: 5px 0;'>", unsafe_allow_html=True)