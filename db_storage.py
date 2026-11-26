def import_data_historis_lengkap(df_excel):
    db = _db()
    count = 0
    
    # 1. Bersihkan nama kolom
    df_excel.columns = df_excel.columns.astype(str).str.strip()
    cols_lower = {c.lower(): c for c in df_excel.columns}

    # 2. Deteksi File
    is_file_nilai = any(k in cols_lower for k in ["pend. agama", "matematika", "b. indo", "b.indonesia", "ipa"])
    is_file_survei_mentah = "saya merasa puas ketika berhasil memahami materi yang sulit" in cols_lower

    # =========================================================
    # PROSES FILE NILAI
    # =========================================================
    if is_file_nilai:
        kolom_nisn = cols_lower.get("nisn") or cols_lower.get("nis")
        if not kolom_nisn: return 0
        
        for _, row in df_excel.iterrows():
            nisn = clean_nisn(row[kolom_nisn])
            if not nisn: continue
            
            data_m = {"NISN": nisn}
            if "nama" in cols_lower: data_m["Nama"] = row[cols_lower["nama"]]
            if "kelas" in cols_lower: data_m["Kelas"] = row[cols_lower["kelas"]]

            nilai_list = []
            for k_low, k_asli in cols_lower.items():
                if k_low in ["no", "nis", "nisn", "nama", "kelas", "rata-rata", "rata-2", "jumlah", "peringkat", "1", "2"]:
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

    # =========================================================
    # PROSES FILE KUISIONER MENTAH (DENGAN FUZZY MATCHING)
    # =========================================================
    elif is_file_survei_mentah:
        import re
        from difflib import SequenceMatcher

        # Fungsi pembersih nama (hapus titik, koma, spasi berlebih)
        def clean_name(n):
            n = str(n).lower()
            n = re.sub(r'[^a-z0-9\s]', '', n) # Hapus simbol
            return " ".join(n.split()) # Hapus spasi ganda
        
        # Ambil semua Nama & NISN yang sudah ada di database
        semua_master = list(db["siswa_master"].find({}, {"NISN": 1, "Nama": 1}))
        db_names = []
        for m in semua_master:
            if m.get("Nama"):
                db_names.append({
                    "NISN": m.get("NISN"), 
                    "Nama_Clean": clean_name(m.get("Nama"))
                })

        def get_val(keyword):
            for k_low, k_asli in cols_lower.items():
                if keyword in k_low: return float(row[k_asli]) if pd.notna(row[k_asli]) else 3.0
            return 3.0

        for _, row in df_excel.iterrows():
            nama_siswa_raw = str(row.get(cols_lower.get("nama", ""), "")).strip()
            if not nama_siswa_raw: continue
            
            nama_siswa_clean = clean_name(nama_siswa_raw)
            
            # --- FUZZY MATCHING NAMA ---
            best_match_nisn = None
            highest_ratio = 0.0
            
            for db_siswa in db_names:
                db_name = db_siswa["Nama_Clean"]
                # Cek 1: Apakah salah satu nama mengandung kata dari nama lain?
                # Cek 2: Hitung kemiripan huruf (SequenceMatcher)
                ratio = SequenceMatcher(None, nama_siswa_clean, db_name).ratio()
                
                # Bonus jika ada kata yang sama persis (misal "septa" ada di kedua nama)
                if any(word in db_name.split() for word in nama_siswa_clean.split() if len(word) > 3):
                    ratio += 0.2

                if ratio > highest_ratio:
                    highest_ratio = ratio
                    best_match_nisn = db_siswa["NISN"]

            # Jika kemiripan di atas 60%, anggap itu orang yang sama
            nisn = best_match_nisn if highest_ratio >= 0.6 else None

            if not nisn: 
                # Jika tetap tidak ketemu, lewati
                continue

            data_s = {"NISN": nisn, "Nama": nama_siswa_raw}
            if "kelas" in cols_lower: data_s["Kelas"] = row[cols_lower["kelas"]]
            
            # --- RUMUS KONSTRUK PSIKOLOGIS (Skala 1-5) ---
            se_cols = ["saya merasa puas ketika berhasil", "saya merasa nyaman dengan teman", "saya yakin bisa memahami materi", "saya yakin bisa mengerjakan tugas", "saya yakin bisa menguasai keterampilan", "saya yakin bisa mendapat nilai bagus"]
            data_s["Self_Efficacy"] = sum([get_val(c) for c in se_cols]) / 6.0

            lm_cols = ["saya senang belajar hal-hal baru", "saya belajar karena ingin tahu", "saya belajar karena ingin mendapat nilai", "saya belajar karena orang tua", "saya merasa belajar di sekolah membuang waktu"]
            lm_vals = [get_val(c) for c in lm_cols]
            lm_vals[4] = 6 - lm_vals[4]
            data_s["Learning_Motivation"] = sum(lm_vals) / 5.0

            ss_cols = ["guru-guru menuntut saya", "guru-guru di sekolah saya mengajar", "guru memberikan perhatian", "suasana di sekolah saya mendukung", "fasilitas sekolah (perpustakaan"]
            data_s["School_Support"] = sum([get_val(c) for c in ss_cols]) / 5.0

            pi_cols = ["orang tua menyediakan waktu", "orang tua mendorong saya", "orang tua menghargai prestasi", "orang tua menyediakan fasilitas", "orang tua memantau kegiatan", "orang tua menanyakan tentang nilai"]
            data_s["Parental_Involvement"] = sum([get_val(c) for c in pi_cols]) / 6.0

            aa_cols = ["saya merasa cemas menghadapi ujian", "saya merasa tertekan dengan target", "saya malas belajar (r)", "menurut saya, materi pelajaran terlalu sulit", "orang tua saya menuntut nilai yang tinggi"]
            aa_vals = [get_val(c) for c in aa_cols]
            aa_vals[2] = 6 - aa_vals[2]
            aa_vals[3] = 6 - aa_vals[3]
            data_s["Academic_Anxiety"] = sum(aa_vals) / 5.0

            # Simpan
            db["siswa_survei"].update_one({"NISN": nisn}, {"$set": data_s}, upsert=True)
            db["siswa_akun"].update_one({"NISN": nisn}, {"$set": {"sudah_isi_survei": True}}, upsert=True)
            count += 1
            
        return count

    # =========================================================
    # PROSES FILE DATA_SIAP_UPLOAD (Excel yang sudah dirata-rata)
    # =========================================================
    else:
        kolom_nisn = cols_lower.get("nisn") or cols_lower.get("nis")
        if not kolom_nisn: return 0

        for _, row in df_excel.iterrows():
            nisn = clean_nisn(row[kolom_nisn])
            if not nisn: continue
            
            data_s = {"NISN": nisn}
            for k_low, k_asli in cols_lower.items():
                if any(x in k_low for x in ["self_efficacy", "academic_anxiety", "school_support", "parental_involvement", "learning_motivation"]):
                    if pd.notna(row[k_asli]): data_s[k_asli] = float(row[k_asli])
                elif any(x in k_low for x in ["usia", "pendidikan terakhir", "pendapatan", "jarak", "fasilitas", "gadget", "belajar di rumah", "les / bimbingan", "ekstrakurikuler"]):
                    if pd.notna(row[k_asli]): data_s[k_asli] = row[k_asli]
            
            for k_low, k_asli in cols_lower.items():
                if k_low in ["nama", "nama siswa"]: data_s["Nama"] = row[k_asli]
                if k_low == "kelas": data_s["Kelas"] = row[k_asli]

            db["siswa_survei"].update_one({"NISN": nisn}, {"$set": data_s}, upsert=True)
            db["siswa_akun"].update_one({"NISN": nisn}, {"$set": {"sudah_isi_survei": True}}, upsert=True)
            count += 1

        return count