import pandas as pd
import secrets
import string
import hashlib

def hash_password(p): 
    return hashlib.sha256(p.encode()).hexdigest()

def generate_password(length=8):
    # Hindari karakter ambigu seperti O, 0, l, 1
    chars = "".join(c for c in string.ascii_letters + string.digits if c not in "0O1lI")
    return "".join(secrets.choice(chars) for _ in range(length))

def proses_excel_dapodik(file):
    """Membaca Excel Dapodik dan membuat password otomatis."""
    # Asumsi kolom Excel: NISN, Nama, Kelas, Nilai
    df = pd.read_excel(file, dtype={"NISN": str})
    df["NISN"] = df["NISN"].str.strip()
    
    # Buat password dan hash
    df["Password"] = [generate_password() for _ in range(len(df))]
    df["Password_Hash"] = df["Password"].apply(hash_password)
    
    return df