# -*- coding: utf-8 -*-
import streamlit as st
import dns.resolver
from datetime import datetime
import pytz
import hashlib
import pandas as pd
from pymongo import MongoClient

# ================================================================
# KONEKSI MONGODB
# ================================================================
dns.resolver.default_resolver = dns.resolver.Resolver(configure=False)
dns.resolver.default_resolver.nameservers = ['8.8.8.8', '8.8.4.4']

@st.cache_resource
def _get_client():
    return MongoClient(st.secrets["MONGODB_URI"], serverSelectionTimeoutMS=30000)

def _db():
    return _get_client()["saa_smpn6"]

def hash_password(p):
    return hashlib.sha256(p.encode()).hexdigest()

def clean_nisn(nisn):
    if pd.isna(nisn): return ""
    return str(nisn).replace('.0', '').strip()

# ================================================================
# MANAJEMEN PENGGUNA (GURU/ADMIN)
# ================================================================
def load_users():
    users = {}
    for doc in _db().users.find():
        username = doc.pop("_id")
        users[username] = doc
    return users

def save_users(users_dict):
    col = _db().users
    col.delete_many({})
    if users_dict:
        docs = [{"_id": u, **data} for u, data in users_dict.items()]
        col.insert_many(docs)

def update_user_password(username, new_hashed_pw):
    _db().users.update_one({"_id": username}, {"$set": {"password": new_hashed_pw}})

def daftar_user_baru(username, nama, password_hashed, role="guru"):
    db = _db()
    if db.users.find_one({"_id": username}):
        return False, "Username sudah digunakan."
    db.users.insert_one({
        "_id": username, "nama": nama, "password": password_hashed,
        "role": role, "status": "pending"
    })
    return True, "Pendaftaran berhasil! Silakan tunggu persetujuan Admin."

# ================================================================
# PORTAL SISWA & SURVEI
# ================================================================
def login_siswa(nisn, password_plain):
    db = _db()
    nisn_clean = clean_nisn(nisn)
    
    akun = db["siswa_akun"].find_one({"NISN": nisn_clean})
    if not akun:
        # AUTO-REGISTER
        if str(password_plain).strip() != nisn_clean:
            return False, "NISN belum terdaftar. Untuk mendaftar mandiri, pastikan Password sama dengan NISN Anda."
            
        db["siswa_akun"].insert_one({
            "NISN": nisn_clean, 
            "Password_Hash": hash_password(nisn_clean), 
            "sudah_isi_survei": False
        })
        return True, "Pendaftaran berhasil! Silakan isi formulir survei."
    else:
        if akun["Password_Hash"] != hash_password(password_plain): 
            return False, "Password salah. Pastikan menggunakan NISN Anda."
        return True, "Login berhasil"

def simpan_survei_siswa(nisn, profil_konstruk):
    db = _db()
    nisn_clean = clean_nisn(nisn)
    wib = datetime.now(pytz.timezone("Asia/Jakarta")).isoformat(timespec="seconds")
    
    db["siswa_survei"].update_one(
        {"NISN": nisn_clean},
        {"$set": {**profil_konstruk, "NISN": nisn_clean, "timestamp": wib}},
        upsert=True
    )
    db["siswa_akun"].update_one(
        {"NISN": nisn_clean}, {"$set": {"sudah_isi_survei": True}}
    )
    return True

def reset_password_siswa(nisn):
    db = _db()
    nisn_clean = clean_nisn(nisn)
    try:
        db["siswa_akun"].update_one(
            {"NISN": nisn_clean},
            {"$set": {"Password_Hash": hash_password(nisn_clean)}}
        )
        return True, "Password berhasil direset menjadi NISN."
    except Exception as e:
        return False, f"Gagal reset: {e}"

def hapus_survei_siswa(nisn):
    """Menghapus TOTAL semua jejak siswa (Akun, Survei, dan Nilai)"""
    db = _db()
    nisn_clean = clean_nisn(nisn)
    try:
        db["siswa_akun"].delete_one({"NISN": nisn_clean})
        db["siswa_survei"].delete_one({"NISN": nisn_clean})
        db["siswa_master"].delete_one({"NISN": nisn_clean}) 
        return True, "Data survei, akun, dan nilai siswa berhasil dihapus total. NISN dapat digunakan kembali."
    except Exception as e:
        return False, f"Gagal menghapus: {e}"

def hapus_nilai_saja(nisn):
    """Menghapus HANYA data nilai (dari siswa_master). Survei tetap aman."""
    db = _db()
    nisn_clean = clean_nisn(nisn)
    try:
        db["siswa_master"].delete_one({"NISN": nisn_clean})
        return True, f"✅ Nilai NISN {nisn_clean} berhasil dihapus. (Survei tetap aman)."
    except Exception as e:
        return False, f"Gagal menghapus nilai: {e}"

# ================================================================
# DATA GABUNGAN (UNTUK DASHBOARD)
# ================================================================
def get_data_gabungan():
    db = _db()
    masters_raw = list(db["siswa_master"].find({}, {"_id": 0}))
    surveis_raw = list(db["siswa_survei"].find({}, {"_id": 0}))
    
    masters = {clean_nisn(m.get("NISN", "")): m for m in masters_raw}
    surveis = {clean_nisn(s.get("NISN", "")): s for s in surveis_raw}
    
    all_nisn = set(masters.keys()) | set(surveis.keys())
    
    data_gabungan = []
    for nisn in all_nisn:
        if not nisn: continue
        m = masters.get(nisn, {})
        s = surveis.get(nisn, {})
        
        row = {**m, **s}
        if "Nama" not in row or not row["Nama"]: 
            row["Nama"] = s.get("Nama", m.get("Nama", f"Siswa {nisn}"))
        if "Kelas" not in row or not row["Kelas"]: 
            row["Kelas"] = s.get("Kelas", m.get("Kelas", "-"))
            
        row["NISN"] = nisn
        row["Nilai Akademik"] = m.get("Nilai", 0)
        data_gabungan.append(row)
        
    # Urutkan berdasarkan NISN terbesar ke terkecil
    data_gabungan.sort(key=lambda x: str(x.get("NISN", "")), reverse=True)
    return data_gabungan

# ================================================================
# IMPORT & MANAJEMEN NILAI (GURU)
# ================================================================
def import_data_historis_lengkap(df_excel):
    db = _db()
    count_master = 0
    
    # 1. Deteksi nama kolom NISN (bisa "NISN" atau "NIS")
    kolom_nisn = "NISN" if "NISN" in df_excel.columns else ("NIS" if "NIS" in df_excel.columns else None)
    if not kolom_nisn:
        return 0 # Gagal, tidak ada NISN/NIS

    # 2. Cek apakah ini file SURVEI (punya kolom Self_Efficacy) atau file NILAI
    is_file_survei = "Self_Efficacy" in df_excel.columns
    is_file_nilai = any(c in df_excel.columns for c in ["Pend. Agama", "Matematika", "B Indo", "IPA"])

    # 3. Daftar kolom yang diabaikan (bukan nilai mapel)
    kolom_diabaikan = [
        "NO", "NIS", "NISN", "Nama", "NAMA", "Nama Siswa", "NAMA SISWA", "Kelas", "KELAS", 
        "Rata-rata", "RATA-RATA", "Rata-2", "Jumlah", "PERINGKAT", "1", "2"
    ] + list(NON_LIKERT_ITEMS.keys() if 'NON_LIKERT_ITEMS' in globals() else []) + [
        "Self_Efficacy", "Academic_Anxiety", "School_Support", "Parental_Involvement", "Learning_Motivation"
    ]
    
    # Cari kolom nilai mapel (yang isinya angka)
    kolom_mapel = []
    for c in df_excel.columns:
        c_clean = str(c).strip()
        if c_clean in kolom_diabaikan: continue
        if pd.api.types.is_numeric_dtype(df_excel[c]):
            kolom_mapel.append(c)

    # 4. Loop setiap baris siswa
    for index, row in df_excel.iterrows():
        nisn = clean_nisn(row[kolom_nisn])
        if not nisn: continue
        
        # ==========================================
        # PROSES DATA (JIKA INI FILE SURVEI)
        # ==========================================
        if is_file_survei:
            data_survei = {"NISN": nisn}
            
            # Ambil psikologis
            for k in ["Self_Efficacy", "Academic_Anxiety", "School_Support", "Parental_Involvement", "Learning_Motivation"]:
                if k in df_excel.columns and pd.notna(row[k]):
                    data_survei[k] = float(row[k])
            
            # Ambil demografi
            if 'NON_LIKERT_ITEMS' in globals():
                for k in NON_LIKERT_ITEMS.keys():
                    if k in df_excel.columns and pd.notna(row[k]):
                        data_survei[k] = row[k]
                        
            if "Nama" in df_excel.columns: data_survei["Nama"] = row["Nama"]
            if "Kelas" in df_excel.columns: data_survei["Kelas"] = row["Kelas"]

            db["siswa_survei"].update_one({"NISN": nisn}, {"$set": data_survei}, upsert=True)
            
            # Tandai akun sudah isi survei
            if not db["siswa_akun"].find_one({"NISN": nisn}):
                db["siswa_akun"].insert_one({
                    "NISN": nisn, 
                    "Password_Hash": hash_password(nisn), 
                    "sudah_isi_survei": True
                })
            else:
                db["siswa_akun"].update_one({"NISN": nisn}, {"$set": {"sudah_isi_survei": True}})
                
            count_master += 1
            
        # ==========================================
        # PROSES DATA (JIKA INI FILE NILAI)
        # ==========================================
        elif is_file_nilai:
            data_master = {"NISN": nisn}
            
            if "NAMA" in df_excel.columns: data_master["Nama"] = row["NAMA"]
            elif "Nama" in df_excel.columns: data_master["Nama"] = row["Nama"]
            elif "Nama Siswa" in df_excel.columns: data_master["Nama"] = row["Nama Siswa"]
            
            if "Kelas" in df_excel.columns: data_master["Kelas"] = row["Kelas"]
            
            # Ambil nilai mapel
            nilai_terkumpul = []
            for m in kolom_mapel:
                val = row[m]
                if pd.notna(val) and isinstance(val, (int, float, float)):
                    data_master[f"Nilai_{m}"] = float(val)
                    nilai_terkumpul.append(float(val))
            
            # Hitung rata-rata
            if nilai_terkumpul:
                data_master["Nilai"] = sum(nilai_terkumpul) / len(nilai_terkumpul)
            else:
                # Kalau Excel sudah punya kolom rata-rata, pakai itu
                if "RATA-RATA" in df_excel.columns and pd.notna(row["RATA-RATA"]):
                    data_master["Nilai"] = float(row["RATA-RATA"])
                elif "Rata-rata" in df_excel.columns and pd.notna(row["Rata-rata"]):
                    data_master["Nilai"] = float(row["Rata-rata"])
                else:
                    data_master["Nilai"] = 0.0
                    
            db["siswa_master"].update_one({"NISN": nisn}, {"$set": data_master}, upsert=True)
            
            # Buatkan akun siswa jika belum ada
            if not db["siswa_akun"].find_one({"NISN": nisn}):
                db["siswa_akun"].insert_one({
                    "NISN": nisn, 
                    "Password_Hash": hash_password(nisn), 
                    "sudah_isi_survei": False
                })
                
            count_master += 1
            
    return count_master

def tambah_siswa_baru(nisn, nama, kelas, nilai):
    db = _db()
    nisn_clean = clean_nisn(nisn)
    db["siswa_master"].update_one(
        {"NISN": nisn_clean},
        {"$set": {"Nama": str(nama).strip(), "Kelas": str(kelas).strip(), "Nilai": float(nilai)}},
        upsert=True
    )
    if not db["siswa_akun"].find_one({"NISN": nisn_clean}):
        db["siswa_akun"].insert_one({"NISN": nisn_clean, "Password_Hash": hash_password(nisn_clean), "sudah_isi_survei": False})
    return True

def update_nilai_siswa(nisn, data_dict):
    db = _db()
    nisn_clean = clean_nisn(nisn)
    try:
        db["siswa_master"].update_one({"NISN": nisn_clean}, {"$set": data_dict}, upsert=True)
        return True, "Data nilai berhasil diperbarui."
    except Exception as e:
        return False, f"Gagal update nilai: {e}"

def get_semua_akun_siswa():
    db = _db()
    masters = list(db["siswa_master"].find({}, {"_id": 0}))
    data = [{"NISN": m.get("NISN"), "Nama": m.get("Nama"), "Kelas": m.get("Kelas"), "Password Default": m.get("NISN")} for m in masters]
    
    df = pd.DataFrame(data)
    if not df.empty:
        df["NISN"] = df["NISN"].astype(str)
        df = df.sort_values(by="NISN", ascending=False)
        
    return df