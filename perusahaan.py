import pandas as pd
from rapidfuzz import process, fuzz

# Fungsi untuk mencocokkan nama perusahaan
def match_perusahaan(data_row, master_list):
    # """
    # Fungsi mencocokkan nama perusahaan dengan master menggunakan rapidfuzz.
    # Ambang batas kemiripan ditetapkan sebagai 70%.
    # """
    best_match = process.extractOne(
        data_row['data'],  # String untuk dicocokkan
        master_list,       # Daftar nama master
        scorer=fuzz.token_sort_ratio  # Gunakan metode token_sort_ratio
    )
    if best_match and best_match[1] >= 70:  # Jika skor >= 70 dianggap cocok
        return best_match[0]  # Return nama master terbaik
    return None  # Jika tidak ada kecocokan

# Buka file Excel
file_path = "data perusahaan.xlsx"  # Ganti dengan path file Anda
data = pd.read_excel(file_path)

# Pastikan kolom 'data' dan 'perusahaan' berbentuk string
data['data'] = data['data'].astype(str)
data['perusahaan'] = data['perusahaan'].astype(str)

# Buat daftar master dari kolom 'perusahaan'
master_list = data['perusahaan'].tolist()

# Cocokkan data dan tambahkan kolom 'kode.1' berdasarkan hasil
data['kode.1'] = data.apply(
    lambda row: match_perusahaan(row, master_list), axis=1
)

# Simpan hasil ke file baru
output_path = "data_perusahaan_matched.xlsx"  # Nama file output
data.to_excel(output_path, index=False)
print(f"File hasil telah disimpan di {output_path}")
