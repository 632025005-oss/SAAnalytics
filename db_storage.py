# -*- coding: utf-8 -*-
import streamlit as st
import dns.resolver
from datetime import datetime
import pytz
from bson.objectid import ObjectId
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

# ================================================================
# MANAJEMEN PENGGUNA (GURU & ADMIN)
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
# MANAJEMEN SISWA (MASTER)
# ================================================================
def tambah_siswa_baru(nisn, nama, kelas, nilai):
    db = _db()
    nisn_str = str(nisn).strip()
    db["siswa_master"].update_one(
        {"NISN": nisn_str},
        {"$set": {"Nama": str(nama), "Kelas": str(kelas), "Nilai": float(nilai)}},
        upsert=True
    )
    if not db["siswa_akun"].find_one({"NISN": nisn_str}):
        db["siswa_akun"].insert_one({
            "NISN": nisn_str, "Password_Hash": hash_password(nisn_str), "sudah_isi_survei": False
        })
    return True

def update_nilai_siswa(nisn, data_dict):
    """Memperbarui data nilai (bisa rata-rata atau per mapel sekaligus)."""
    db = _db()
    try:
        db["siswa_master"].update_one(
            {"NISN": str(nisn).strip()},
            {"$set": data_dict}
        )
        return True, "Data nilai berhasil diperbarui."
    except Exception as e:
        return False, f"Gagal update nilai: {e}"

def get_semua_akun_siswa():
    db = _db()
    masters = list(db["siswa_master"].find({}, {"_id": 0}))
    data = [{"NISN": m["NISN"], "Nama": m["Nama"], "Kelas": m["Kelas"], "Password Default": m["NISN"]} for m in masters]
    return pd.DataFrame(data)

# ================================================================
# SURVEI & AKUN SISWA
# ================================================================
def login_siswa(nisn, password_plain):
    db = _db()
    akun = db["siswa_akun"].find_one({"NISN": str(nisn).strip()})
    if not akun: return False, "NISN tidak terdaftar."
    if akun["Password_Hash"] != hash_password(password_plain): return False, "Password salah."
    return True, "Login berhasil"

def simpan_survei_siswa(nisn, profil_konstruk):
    db = _db()
    wib = datetime.now(pytz.timezone("Asia/Jakarta")).isoformat(timespec="seconds")
    db["siswa_survei"].update_one(
        {"NISN": str(nisn).strip()},
        {"$set": {**profil_konstruk, "NISN": str(nisn).strip(), "timestamp": wib}},
        upsert=True
    )
    db["siswa_akun"].update_one(
        {"NISN": str(nisn).strip()}, {"$set": {"sudah_isi_survei": True}}
    )
    return True

# ================================================================
# DATA GABUNGAN (KUNCI DASHBOARD)
# ================================================================
def get_data_gabungan():
    db = _db()
    masters = list(db["siswa_master"].find({}, {"_id": 0}))
    # Paksa NISN jadi string agar sinkron
    surveis = {str(s["NISN"]).strip(): s for s in db["siswa_survei"].find({}, {"_id": 0})}
    
    data_gabungan = []
    for m in masters:
        nisn = str(m["NISN"]).strip()
        if nisn in surveis: 
            row = {**m, **surveis[nisn]}
            row["Nilai Akademik"] = row.pop("Nilai", 0) # Ubah nama 'Nilai' jadi 'Nilai Akademik'
            data_gabungan.append(row)
        else:
            row = {**m}
            row["Nilai Akademik"] = row.pop("Nilai", 0)
            data_gabungan.append(row)
    return data_gabungan

# ================================================================
# IMPORT EXCEL (HISTORIS LENGKAP)
# ================================================================
def import_data_historis_lengkap(df_excel):
    db = _db()
    count_master = 0
    count_survei = 0
    
    # Kolom yang masuk ke Master (Biodata & Nilai Mapel)
    kolom_master = ["NISN", "Nama", "Kelas", "Nilai"] + [c for c in df_excel.columns if '_S1' in c or '_S2' in c]
    # Kolom yang masuk ke Survei (Psikologis & Demografi)
    kolom_survei = [
        "Usia", "Pendidikan Terakhir Ayah", "Pendidikan Terakhir Ibu", 
        "Perkiraan Pendapatan Orang Tua per bulan", "Jarak rumah ke sekolah", 
        "Fasilitas yang tersedia di rumah (centang semua yang sesuai):", 
        "Rata-rata berapa jam per hari kamu menggunakan gadget/HP di luar keperluan belajar (bermain game, media sosial, menonton video)?", 
        "Rata-rata berapa jam kamu BELAJAR DI RUMAH setiap hari (di luar jam sekolah)?", 
        "Berapa jam per MINGGU kamu mengikuti LES / BIMBINGAN BELAJAR di luar sekolah  ?", 
        "Berapa banyak kegiatan EKSTRAKURIKULER yang kamu ikuti di sekolah selama kelas 9 semester 1?", 
        "Self_Efficacy", "Academic_Anxiety", "School_Support", "Parental_Involvement", "Learning_Motivation"
    ]
    
    for index, row in df_excel.iterrows():
        nisn = str(row["NISN"]).replace('.0', '').strip()
        
        # 1. Simpan ke Master
        data_master = {"NISN": nisn}
        for k in kolom_master:
            if k != "NISN" and k in df_excel.columns: 
                data_master[k] = row[k]
        db["siswa_master"].update_one({"NISN": nisn}, {"$set": data_master}, upsert=True)
        count_master += 1
        
        # 2. Simpan ke Survei
        if all(k in df_excel.columns for k in kolom_survei):
            data_survei = {"NISN": nisn, "timestamp": datetime.now().isoformat()}
            for k in kolom_survei:
                val = row[k] if pd.notna(row[k]) else (3.0 if k in ["Self_Efficacy", "Academic_Anxiety", "School_Support", "Parental_Involvement", "Learning_Motivation"] else 0)
                data_survei[k] = float(val) if isinstance(val, (int, float)) else val
            db["siswa_survei"].update_one({"NISN": nisn}, {"$set": data_survei}, upsert=True)
            count_survei += 1
            
        # 3. Buatkan Akun
        if not db["siswa_akun"].find_one({"NISN": nisn}):
            db["siswa_akun"].insert_one({"NISN": nisn, "Password_Hash": hash_password(nisn), "sudah_isi_survei": True})
            
    return count_master, count_survei

def import_dapodik_ke_db(df_dapodik):
    db = _db()
    db["siswa_master"].delete_many({})
    db["siswa_akun"].delete_many({})
    master_records = df_dapodik[["NISN", "Nama", "Kelas", "Nilai"]].to_dict("records")
    if master_records: db["siswa_master"].insert_many(master_records)
    return len(master_records)

def tambah_log(entry):
    _db().log.insert_one(dict(entry))