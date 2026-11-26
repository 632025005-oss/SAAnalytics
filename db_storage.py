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
        db["siswa_master"].delete_one({"NISN": nisn_clean}) 
        return True, "Data survei dan akun siswa berhasil dihapus."
    except Exception as e:
        return False, f"Gagal menghapus: {e}"

def hapus_nilai_saja(nisn):
    db = _db()
    nisn_clean = clean_nisn(nisn)
    try:
        db["siswa_master"].delete_one({"NISN": nisn_clean})
        return True, f"✅ Nilai NISN {nisn_clean} berhasil dihapus. (Survei tetap aman)."
    except Exception as e:
        return False, f"Gagal menghapus nilai: {e}"

# ================================================================
# DATA GABUNGAN
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
        
        # Cek eksplisit apakah data survei (Self_Efficacy) ada isinya
        if "Self_Efficacy" in s and s["Self_Efficacy"] is not None:
            row["Self_Efficacy"] = s["Self_Efficacy"]
            
        data_gabungan.append(row)
        
    data_gabungan.sort(key=lambda x: str(x.get("NISN", "")), reverse=True)
    return data_gabungan

# ================================================================
# IMPORT EXCEL (ANTI GAGAL)
# ================================================================
def import_data_historis_lengkap(df_excel):
    db = _db()
    count = 0
    
    # Normalisasi nama kolom (hapus spasi di ujung, samakan huruf kecil)
    df_excel.columns = df_excel.columns.astype(str).str.strip()
    cols_lower = {c.lower(): c for c in df_excel.columns}

    # Cari kolom NISN / NIS
    kolom_nisn = None
    for k in ["nisn", "nis"]:
        if k in cols_lower: kolom_nisn = cols_lower[k]; break
    if not kolom_nisn: return 0

    # Deteksi tipe file
    is_survei = any(k in cols_lower for k in ["self_efficacy", "self efficacy", "academic_anxiety"])
    is_nilai = any(k in cols_lower for k in ["pend. agama", "matematika", "ipa", "b indo"])

    for _, row in df_excel.iterrows():
        nisn = clean_nisn(row[kolom_nisn])
        if not nisn: continue
        
        # PROSES SURVEI
        if is_survei:
            data_s = {"NISN": nisn}
            for k_low, k_asli in cols_lower.items():
                # Tangkap semua kolom survei psikologis
                if any(x in k_low for x in ["self_efficacy", "academic_anxiety", "school_support", "parental_involvement", "learning_motivation"]):
                    if pd.notna(row[k_asli]): data_s[k_asli] = float(row[k_asli])
                # Tangkap kolom demografi
                elif any(x in k_low for x in ["usia", "pendidikan terakhir", "pendapatan", "jarak", "fasilitas", "gadget", "belajar di rumah", "les / bimbingan", "ekstrakurikuler"]):
                    if pd.notna(row[k_asli]): data_s[k_asli] = row[k_asli]
            
            # Tangkap Nama & Kelas
            for k_low, k_asli in cols_lower.items():
                if k_low in ["nama", "nama siswa"]: data_s["Nama"] = row[k_asli]
                if k_low == "kelas": data_s["Kelas"] = row[k_asli]

            db["siswa_survei"].update_one({"NISN": nisn}, {"$set": data_s}, upsert=True)
            db["siswa_akun"].update_one({"NISN": nisn}, {"$set": {"sudah_isi_survei": True}}, upsert=True)
            count += 1

        # PROSES NILAI
        elif is_nilai:
            data_m = {"NISN": nisn}
            for k_low, k_asli in cols_lower.items():
                if k_low in ["nama", "nama siswa"]: data_m["Nama"] = row[k_asli]
                if k_low == "kelas": data_m["Kelas"] = row[k_asli]

            # Ambil semua angka mapel
            nilai_list = []
            for k_low, k_asli in cols_lower.items():
                if k_low in ["no", "nis", "nisn", "nama", "nama siswa", "kelas", "rata-rata", "rata-2", "jumlah", "peringkat", "1", "2"]:
                    continue
                if pd.api.types.is_numeric_dtype(df_excel[k_asli]) and pd.notna(row[k_asli]):
                    data_m[f"Nilai_{k_asli}"] = float(row[k_asli])
                    nilai_list.append(float(row[k_asli]))
            
            data_m["Nilai"] = sum(nilai_list) / len(nilai_list) if nilai_list else 0.0
            db["siswa_master"].update_one({"NISN": nisn}, {"$set": data_m}, upsert=True)
            
            if not db["siswa_akun"].find_one({"NISN": nisn}):
                db["siswa_akun"].insert_one({"NISN": nisn, "Password_Hash": hash_password(nisn), "sudah_isi_survei": False})
            count += 1

    return count

def tambah_siswa_baru(nisn, nama, kelas, nilai):
    db = _db()
    nisn_clean = clean_nisn(nisn)
    db["siswa_master"].update_one({"NISN": nisn_clean}, {"$set": {"Nama": str(nama).strip(), "Kelas": str(kelas).strip(), "Nilai": float(nilai)}}, upsert=True)
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