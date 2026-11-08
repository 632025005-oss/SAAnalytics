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
# MANAJEMEN PENGGUNA
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
    db = _db()
    nisn_clean = clean_nisn(nisn)
    try:
        db["siswa_akun"].delete_one({"NISN": nisn_clean})
        db["siswa_survei"].delete_one({"NISN": nisn_clean})
        return True, "Data survei dan akun siswa berhasil dihapus. NISN dapat digunakan kembali."
    except Exception as e:
        return False, f"Gagal menghapus: {e}"

# ================================================================
# DATA GABUNGAN (DENGAN FITUR SORTING NISN)
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
        if "Nama" not in row or not row["Nama"]: row["Nama"] = s.get("Nama", f"Siswa {nisn}")
        if "Kelas" not in row or not row["Kelas"]: row["Kelas"] = s.get("Kelas", "-")
        row["NISN"] = nisn
        row["Nilai Akademik"] = m.get("Nilai", 0)
        data_gabungan.append(row)
        
    # TAMBAHAN: Urutkan berdasarkan NISN terbesar ke terkecil
    data_gabungan.sort(key=lambda x: str(x.get("NISN", "")), reverse=True)
        
    return data_gabungan

# ================================================================
# IMPORT & MANAJEMEN NILAI (GURU)
# ================================================================
def import_data_historis_lengkap(df_excel):
    db = _db()
    count_master = 0
    kolom_master = ["NISN", "Nama", "Kelas", "Nilai"] + [c for c in df_excel.columns if '_S1' in c or '_S2' in c]
    
    for index, row in df_excel.iterrows():
        nisn = clean_nisn(row["NISN"])
        if not nisn: continue
        
        data_master = {"NISN": nisn}
        for k in kolom_master:
            if k != "NISN" and k in df_excel.columns: 
                data_master[k] = row[k]
        db["siswa_master"].update_one({"NISN": nisn}, {"$set": data_master}, upsert=True)
        count_master += 1
        
        if not db["siswa_akun"].find_one({"NISN": nisn}):
            db["siswa_akun"].insert_one({"NISN": nisn, "Password_Hash": hash_password(nisn), "sudah_isi_survei": False})
            
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
    
    # TAMBAHAN: Sortir dataframe berdasarkan NISN
    df = pd.DataFrame(data)
    if not df.empty:
        df["NISN"] = df["NISN"].astype(str)
        df = df.sort_values(by="NISN", ascending=False)
        
    return df