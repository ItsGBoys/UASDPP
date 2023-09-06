import openpyxl

# Buka file Excel
wb = openpyxl.load_workbook('nama_file.xlsx')  # Ganti 'nama_file.xlsx' dengan nama file Anda
sheet = wb.active  # Ini akan memilih sheet pertama, Anda dapat menggantinya sesuai kebutuhan

# Kolom yang ingin Anda periksa (misalnya, kolom A)
kolom_yang_diperiksa = sheet['A']

# Membuat kamus untuk melacak nomor dan barisnya
nomor_dan_baris = {}

for baris, cell in enumerate(kolom_yang_diperiksa, start=1):
    nomor = cell.value
    if nomor is not None:
        if nomor in nomor_dan_baris:
            nomor_dan_baris[nomor].append(baris)
        else:
            nomor_dan_baris[nomor] = [baris]

# Cetak nomor yang sama dan barisnya
for nomor, baris in nomor_dan_baris.items():
    if len(baris) > 1:
        print(f"Nomor '{nomor}' ditemukan di kolom A pada baris ke-{', '.join(map(str, baris))}.")
