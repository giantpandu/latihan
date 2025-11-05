# =========================================================
# Analisis Kelulusan dan Predikat Wisuda Mahasiswa
# Menggunakan pandas dan matplotlib
# =========================================================

import pandas as pd
import matplotlib.pyplot as plt

# === 1. IMPORT DATA ===
# Pastikan file 'data_wisuda.xlsx' ada di folder yang sama
data = pd.read_excel("Data Wisudawan Kel 3 fix.xlsx")

# === 2. PEMBERSIHAN DATA ===
data = data.drop_duplicates(subset=["NIM"])  # hapus duplikat NIM
data = data.dropna(subset=["NIM", "Nama Mahasiswa", "Program Studi", "IPK"])  # hapus baris kosong penting
data["IPK"] = data["IPK"].astype(float)
data["Lama Studi (Semester)"] = data["Lama Studi (Semester)"].astype(int)

# === 3. KLASIFIKASI GRADE BERDASARKAN IPK ===
def get_grade(ipk):
    if 3.75 <= ipk <= 4.00:
        return "A"
    elif 3.50 <= ipk < 3.75:
        return "B+"
    elif 3.00 <= ipk < 3.50:
        return "B"
    elif 2.50 <= ipk < 3.00:
        return "C"
    else:
        return "D"

data["Grade"] = data["IPK"].apply(get_grade)

# === 4. KLASIFIKASI PREdIKAT WISUDA ===
def get_predikat(ipk, lama):
    if ipk >= 3.75 and lama <= 8:
        return "Cumlaude"
    elif ipk >= 3.50 and lama <= 9:
        return "Sangat Memuaskan"
    elif ipk >= 3.00:
        return "Memuaskan"
    else:
        return "Cukup"

data["Predikat Wisuda"] = data.apply(lambda x: get_predikat(x["IPK"], x["Lama Studi (Semester)"]), axis=1)

# === 5. JUMLAH WISUDAWAN PER PRODI ===
jumlah_per_prodi = data["Program Studi"].value_counts()
print("Jumlah Wisudawan per Program Studi:\n", jumlah_per_prodi, "\n")

# === 6. RATA-RATA IPK PER PRODI ===
rata_ipk = data.groupby("Program Studi")["IPK"].mean().round(2)
print("Rata-rata IPK per Program Studi:\n", rata_ipk, "\n")

# === 7. VISUALISASI DATA ===

# Grafik batang jumlah wisudawan per prodi
plt.figure(figsize=(8,5))
jumlah_per_prodi.plot(kind="bar", color="purple")
plt.title("Jumlah Wisudawan per Program Studi")
plt.xlabel("Program Studi")
plt.ylabel("Jumlah Wisudawan")
plt.xticks(rotation=0)
plt.tight_layout()
plt.show()

# Grafik pie predikat wisuda
plt.figure(figsize=(6,6))
data["Predikat Wisuda"].value_counts().plot(kind="pie", autopct="%1.1f%%", startangle=90)
plt.title("Distribusi Predikat Kelulusan")
plt.ylabel("")
plt.tight_layout()
plt.show()

# Grafik tambahan: rata-rata IPK per prodi
plt.figure(figsize=(8,5))
rata_ipk.plot(kind="bar", color="lightgreen")
plt.title("Rata-rata IPK per Program Studi")
plt.xlabel("Program Studi")
plt.ylabel("Rata-rata IPK")
plt.xticks(rotation=0)
plt.tight_layout()
plt.show()

# === 8. SIMPAN HASIL ===
data.to_excel("rekap_wisuda_final.xlsx", index=False)
print("✅ Analisis selesai! File hasil tersimpan: rekap_wisuda_final.xlsx")
