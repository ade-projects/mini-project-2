# =====================================================
# PROGRAM ANALISIS DATA WISUDA
# =====================================================

import pandas as pd
import os
import matplotlib.pyplot as plt

# 1️⃣ Membaca data (pastikan file ada di folder yang sama)
file_path = "Data Wisudawan.xlsx"
df = pd.read_excel(file_path)

# 2️⃣ Fungsi menghitung grade dan predikat
def hitung_grade_predikat(ipk, lama_studi):
    if ipk >= 3.75 and lama_studi <= 8:
        return "A", "Cumlaude"
    elif ipk >= 3.50:
        return "B+", "Sangat Memuaskan"
    elif ipk >= 3.00:
        return "B", "Memuaskan"
    elif ipk >= 2.50:
        return "C", "Cukup"
    else:
        return "D", "Cukup"

# 3️⃣ Terapkan fungsi ke setiap baris data
df[["Grade", "Predikat"]] = df.apply(
    lambda row: hitung_grade_predikat(row["IPK"], row["Lama Studi (Semester)"]),
    axis=1, result_type="expand"
)

# 4️⃣ Tampilkan hasil ke terminal (rata kiri dan rapi)
print("\n=== HASIL ANALISIS DATA WISUDA ===")
print(df.to_string(index=False, justify="left"))

# 5️⃣ Simpan hasil ke file Excel baru (otomatis di folder yang sama)
folder_sekarang = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(folder_sekarang, "rekap_wisuda_final.xlsx")
df.to_excel(output_path, index=False)

print(f"\n✅ File hasil telah disimpan otomatis sebagai: {output_path}")

# 6️⃣ Grafik jumlah wisudawan per prodi
jumlah_per_prodi = df["Program Studi"].value_counts()
plt.figure(figsize=(10, 6))  # Perbesar ukuran untuk keterbacaan yang lebih baik
jumlah_per_prodi.plot(kind="bar", color="skyblue", edgecolor="black")
plt.title("Jumlah Wisudawan per Prodi", fontsize=14, fontweight='bold')
plt.xlabel("Program Studi", fontsize=12)
plt.ylabel("Jumlah Wisudawan", fontsize=12)
plt.xticks(rotation=45, ha="right", fontsize=10)
plt.yticks(fontsize=10)
plt.tight_layout()
plt.grid(axis="y", linestyle="--", alpha=0.7)
# Simpan grafik sebagai file PNG
grafik_bar_path = os.path.join(folder_sekarang, "Grafik_Jumlah_Wisudawan_per_Prodi.png")
plt.savefig(grafik_bar_path, dpi=300, bbox_inches='tight')
print(f"\n✅ Grafik bar telah disimpan sebagai: {grafik_bar_path}")
plt.show()

# 7️⃣ Grafik distribusi predikat wisuda
distribusi_predikat = df["Predikat"].value_counts()
plt.figure(figsize=(8, 8))  # Perbesar ukuran untuk pie chart yang lebih jelas
distribusi_predikat.plot(kind="pie", autopct="%1.1f%%", startangle=90, colors=plt.cm.Pastel1.colors, textprops={'fontsize': 10})
plt.title("Distribusi Predikat Wisuda", fontsize=14, fontweight='bold')
plt.ylabel("")  # Menghilangkan label sumbu Y
plt.tight_layout()
# Simpan grafik sebagai file PNG
grafik_pie_path = os.path.join(folder_sekarang, "Grafik_Distribusi_Predikat_Wisuda.png")
plt.savefig(grafik_pie_path, dpi=300, bbox_inches='tight')
print(f"\n✅ Grafik pie telah disimpan sebagai: {grafik_pie_path}")
plt.show()

# 8️⃣ Grafik rata-rata IPK antar prodi
rata2_ipk_per_prodi = df.groupby("Program Studi")["IPK"].mean().sort_values()
plt.figure(figsize=(10, 6))
rata2_ipk_per_prodi.plot(kind="barh", color="lightgreen", edgecolor="black")
plt.title("Rata-rata IPK per Program Studi", fontsize=14, fontweight='bold')
plt.xlabel("Rata-rata IPK", fontsize=12)
plt.ylabel("Program Studi", fontsize=12)
plt.grid(axis="x", linestyle="--", alpha=0.7)
plt.tight_layout()

# Simpan grafik
grafik_ipk_path = os.path.join(folder_sekarang, "Grafik_RataRata_IPK_per_Prodi.png")
plt.savefig(grafik_ipk_path, dpi=300, bbox_inches='tight')
print(f"\n✅ Grafik IPK rata-rata per prodi disimpan sebagai: {grafik_ipk_path}")
plt.show()

# 9️⃣ Scatter plot hubungan lama studi dan IPK
plt.figure(figsize=(8, 6))
plt.scatter(df["Lama Studi (Semester)"], df["IPK"], color="salmon", edgecolor="black", alpha=0.7)
plt.title("Hubungan Lama Studi dan IPK", fontsize=14, fontweight='bold')
plt.xlabel("Lama Studi (Semester)", fontsize=12)
plt.ylabel("IPK", fontsize=12)
plt.grid(True, linestyle="--", alpha=0.6)
plt.tight_layout()

# Simpan grafik
grafik_scatter_path = os.path.join(folder_sekarang, "Grafik_LamaStudi_vs_IPK.png")
plt.savefig(grafik_scatter_path, dpi=300, bbox_inches='tight')
print(f"\n✅ Grafik hubungan lama studi dan IPK disimpan sebagai: {grafik_scatter_path}")
plt.show()
