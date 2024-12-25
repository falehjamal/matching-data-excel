import pandas as pd
from rapidfuzz import process, fuzz

# Fungsi mencocokkan group berdasarkan kemiripan dengan parameter Nama di Sheet2
def match_group_by_name(sheet2_row, sheet1_groups):
    """
    Mencocokkan Group di Sheet2 dengan Group di Sheet1 berdasarkan kemiripan Nama.
    """
    # Cari kecocokan terbaik untuk Nama terhadap daftar Group di Sheet1
    best_match = process.extractOne(
        sheet2_row['Nama'],  # Nama dari Sheet2
        sheet1_groups,       # Daftar Group dari Sheet1
        scorer=fuzz.token_sort_ratio
    )
    
    # Jika skor kecocokan >= 70, gunakan hasil terbaik
    matched_group = best_match[0] if best_match and best_match[1] >= 70 else None
    
    return matched_group

# Buka file Excel
file_path = "data_group.xlsx"  # Ganti dengan nama file Anda
sheet1 = pd.read_excel(file_path, sheet_name="Sheet1")
sheet2 = pd.read_excel(file_path, sheet_name="Sheet2")

# Pastikan kolom yang digunakan adalah string
sheet1['Group'] = sheet1['Group'].astype(str)
sheet2['Nama'] = sheet2['Nama'].astype(str)

# Ambil daftar Group dari Sheet1
sheet1_groups = sheet1['Group'].tolist()

# Lakukan pencocokan untuk setiap baris di Sheet2
sheet2['Matched Group'] = sheet2.apply(
    lambda row: match_group_by_name(row, sheet1_groups), axis=1
)

# Gabungkan hasil pencocokan dengan kode dari Sheet1
sheet2 = sheet2.merge(sheet1, how='left', left_on='Matched Group', right_on='Group')

# Simpan hasil ke file baru
output_path = "data_group_matched.xlsx"  # Nama file output
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    sheet1.to_excel(writer, sheet_name="Sheet1", index=False)
    sheet2.to_excel(writer, sheet_name="Sheet2", index=False)

print(f"File hasil telah disimpan di {output_path}")
