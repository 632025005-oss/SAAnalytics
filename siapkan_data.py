import pandas as pd
import numpy as np
import difflib

print("Memproses data lengkap (Nilai Mapel + 37 Survei)...")

# 1. Baca Data Nilai (Header ada di baris pertama)
df_nilai = pd.read_excel("Data Nilai kls 9efgh 2026.xlsx", sheet_name="nilai sem 1 - 6", header=0)
# Hapus baris yang kosong atau bukan angka di kolom NIS
df_nilai = df_nilai.dropna(subset=[df_nilai.columns[1]]) # Kolom ke-2 (index 1) adalah NIS

# Buat dataframe baru untuk data yang dibutuhkan
df_nilai_clean = pd.DataFrame()
df_nilai_clean['NISN'] = df_nilai.iloc[:, 1].astype(str).str.replace('.0', '', regex=False).str.strip()
df_nilai_clean['Nama_Nilai'] = df_nilai.iloc[:, 2].astype(str).str.strip()
df_nilai_clean['Nama_Clean'] = df_nilai_clean['Nama_Nilai'].str.lower().str.strip()

# Ambil nilai per mapel (Sem 1 & Sem 2) berdasarkan indeks kolom
mapel_indices = {
    "Pend. Agama": (3, 4), "PPKn": (6, 7), "B Indo": (9, 10), 
    "Matematika": (12, 13), "B Ingg": (15, 16), "Seni Budaya/Prakarya": (18, 19),
    "PJOK": (21, 22), "B Jawa": (24, 25), "IPA": (27, 28), 
    "IPS": (30, 31), "Informatika": (33, 34)
}

for mapel, (col1, col2) in mapel_indices.items():
    df_nilai_clean[f"{mapel}_S1"] = pd.to_numeric(df_nilai.iloc[:, col1], errors='coerce').fillna(80.0)
    df_nilai_clean[f"{mapel}_S2"] = pd.to_numeric(df_nilai.iloc[:, col2], errors='coerce').fillna(80.0)

# Ambil Rata-rata (kolom terakhir, index 36)
df_nilai_clean['Nilai'] = pd.to_numeric(df_nilai.iloc[:, 36], errors='coerce').fillna(80.0)

# 2. Baca Kuesioner
df_kues = pd.read_excel("Kuisioner Thesis (Responses)-asli.xlsx")
df_kues.columns = df_kues.columns.astype(str).str.strip()
df_kues['Nama_Clean'] = df_kues['Nama'].astype(str).str.lower().str.strip()

# Fungsi hitung 5 konstruk
def hitung_konstruk(row):
    se = np.mean([row['Saya merasa puas ketika berhasil memahami materi yang sulit'], row['Saya yakin bisa memahami materi pelajaran yang paling sulit sekalipun'], row['Saya yakin bisa mengerjakan tugas dengan baik jika saya berusaha keras'], row['Saya yakin bisa menguasai keterampilan yang diajarkan di kelas'], row['Saya yakin bisa mendapat nilai bagus jika saya belajar sungguh-sungguh'], row['Saya merasa nyaman dengan teman-teman sekelas']])
    aa = np.mean([6 - row['Saya malas belajar (R)'], row['Saya merasa cemas menghadapi ujian akhir'], row['Saya merasa tertekan dengan target nilai yang harus saya capai'], 6 - row['Menurut saya, materi pelajaran terlalu sulit untuk dipahami (R)']])
    ss = np.mean([row['Guru-guru menuntut saya untuk lulus dengan baik'], row['Guru-guru di sekolah saya mengajar dengan baik'], row['Fasilitas sekolah (perpustakaan, laboratorium) mendukung pembelajaran'], row['Guru memberikan perhatian jika saya kesulitan belajar'], row['Suasana di sekolah saya mendukung untuk belajar']])
    pi = np.mean([row['Orang tua menyediakan waktu untuk membantu saya belajar'], row['Orang tua mendorong saya untuk meraih cita-cita'], row['Orang tua menghargai prestasi yang saya capai'], row['Orang tua menyediakan fasilitas yang saya butuhkan untuk belajar'], row['Orang tua memantau kegiatan belajar saya di rumah'], row['Orang tua menanyakan tentang nilai atau tugas sekolah'], row['Orang tua saya menuntut nilai yang tinggi']])
    lm = np.mean([row['Saya senang belajar hal-hal baru di sekolah'], row['Saya belajar karena ingin tahu lebih banyak'], row['Saya belajar karena ingin mendapat nilai yang bagus'], row['Saya belajar karena orang tua saya menyuruh'], 6 - row['Saya merasa belajar di sekolah membuang waktu']])
    return pd.Series({'Self_Efficacy': se, 'Academic_Anxiety': aa, 'School_Support': ss, 'Parental_Involvement': pi, 'Learning_Motivation': lm})

skor_konstruk = df_kues.apply(hitung_konstruk, axis=1)
non_likert = [
    "Usia", 
    "Pendidikan Terakhir Ayah", 
    "Pendidikan Terakhir Ibu", 
    "Perkiraan Pendapatan Orang Tua per bulan", 
    "Jarak rumah ke sekolah", 
    "Fasilitas yang tersedia di rumah (centang semua yang sesuai):", 
    "Rata-rata berapa jam per hari kamu menggunakan gadget/HP di luar keperluan belajar (bermain game, media sosial, menonton video)?", 
    "Rata-rata berapa jam kamu BELAJAR DI RUMAH setiap hari (di luar jam sekolah)?", 
    "Berapa jam per MINGGU kamu mengikuti LES / BIMBINGAN BELAJAR di luar sekolah  ?", 
    "Berapa banyak kegiatan EKSTRAKURIKULER yang kamu ikuti di sekolah selama kelas 9 semester 1?"
]

df_kues_final = pd.concat([df_kues[['Nama_Clean', 'Kelas'] + non_likert], skor_konstruk], axis=1)

# 3. Gabungkan (Fuzzy Matching)
df_final = pd.merge(df_nilai_clean, df_kues_final, on='Nama_Clean', how='inner')
nama_cocok = set(df_final['Nama_Clean'])

sisa_nilai = df_nilai_clean[~df_nilai_clean['Nama_Clean'].isin(nama_cocok)]
sisa_kues = df_kues_final[~df_kues_final['Nama_Clean'].isin(nama_cocok)]
list_kues = sisa_kues['Nama_Clean'].tolist()
hasil_fuzzy = []

print("Memproses pencocokan nama yang mirip...")
for idx, row in sisa_nilai.iterrows():
    matches = difflib.get_close_matches(row['Nama_Clean'], list_kues, n=1, cutoff=0.6)
    if matches:
        nama_mirip = matches[0]
        kues_mirip = sisa_kues[sisa_kues['Nama_Clean'] == nama_mirip].iloc[0].to_dict()
        
        data_baru = {'NISN': row['NISN'], 'Nama_Nilai': row['Nama_Nilai'], 'Nama_Clean': row['Nama_Clean']}
        for col in df_nilai_clean.columns:
            if col not in ['NISN', 'Nama_Nilai', 'Nama_Clean']:
                data_baru[col] = row[col]
        for col in df_kues_final.columns:
            if col not in ['Nama_Clean']:
                data_baru[col] = kues_mirip[col]
                
        hasil_fuzzy.append(data_baru)
        list_kues.remove(nama_mirip)

if hasil_fuzzy:
    df_final = pd.concat([df_final, pd.DataFrame(hasil_fuzzy)], ignore_index=True)

# 4. Rapikan kolom
kolom_awal = ['NISN', 'Nama_Nilai', 'Kelas', 'Nilai']
kolom_survei = non_likert + ['Self_Efficacy', 'Academic_Anxiety', 'School_Support', 'Parental_Involvement', 'Learning_Motivation']
kolom_mapel = [c for c in df_final.columns if '_S1' in c or '_S2' in c]

df_final = df_final[kolom_awal + kolom_survei + kolom_mapel]
df_final.rename(columns={'Nama_Nilai': 'Nama'}, inplace=True)

df_final.to_excel("Data_Siap_Upload_v2.xlsx", index=False)
print(f"\n✅ Berhasil! {len(df_final)} data siswa lengkap siap di-upload.")