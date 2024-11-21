import requests
import json
from openpyxl import Workbook

def get_data_from_api(url):
    response = requests.get(url)
    if response.status_code == 200:
        try:
            # Mengonversi string menjadi JSON list
            data = json.loads(response.text)
            return data  # Langsung kembalikan data sebagai list
        except json.JSONDecodeError:
            print("Error decoding JSON. Periksa format respons API.")
            return []
    else:
        print(f"Failed to fetch data. Status code: {response.status_code}")
        return []

# Buat workbook untuk Excel
wb = Workbook()
ws = wb.active
ws.title = "Data Sekolah Aceh"

# Buat header dengan kolom tambahan untuk Kabupaten dan Kecamatan
headers = ["No", "Kabupaten", "Kecamatan", "Nama Sekolah", "NPSN", "BP", "Status", "Last Sync", 
           "Jml Sync", "PD", "Rombel", "Guru", "Pegawai", "R. Kelas", "R. Lab", "R. Perpus"]
ws.append(headers)

# 1. Ambil data kabupaten di Provinsi Aceh
kabupaten_url = 'https://dapo.kemdikbud.go.id/rekap/dataSekolah?id_level_wilayah=1&kode_wilayah=060000&semester_id=20241'
kabupaten_data = get_data_from_api(kabupaten_url)
print(f"Jumlah kabupaten: {len(kabupaten_data)}")  # Debugging: Cek jumlah kabupaten

# 2. Ambil data kecamatan untuk setiap kabupaten di Aceh
for kabupaten in kabupaten_data:
    nama_kabupaten = kabupaten['nama']
    kode_wilayah_kabupaten = kabupaten['kode_wilayah'].strip()
    kecamatan_url = f'https://dapo.kemdikbud.go.id/rekap/dataSekolah?id_level_wilayah=2&kode_wilayah={kode_wilayah_kabupaten}&semester_id=20241'
    kecamatan_data = get_data_from_api(kecamatan_url)
    
    # 3. Ambil data sekolah untuk setiap kecamatan dan tambahkan ke sheet yang sama
    for kecamatan in kecamatan_data:
        nama_kecamatan = kecamatan['nama']
        kode_wilayah_kecamatan = kecamatan['kode_wilayah'].strip()
        sekolah_url = f'https://dapo.kemdikbud.go.id/rekap/progresSP?id_level_wilayah=3&kode_wilayah={kode_wilayah_kecamatan}&semester_id=20241&bentuk_pendidikan_id='
        sekolah_data = get_data_from_api(sekolah_url)
        
        # Tambahkan data sekolah ke sheet utama
        if sekolah_data:  # Cek apakah sekolah_data berisi data
            for index, sekolah in enumerate(sekolah_data, start=1):
                row = [
                    index,
                    nama_kabupaten,
                    nama_kecamatan,
                    sekolah.get("nama"),
                    sekolah.get("npsn"),
                    sekolah.get("bentuk_pendidikan"),
                    sekolah.get("status_sekolah"),
                    sekolah.get("sinkron_terakhir"),
                    sekolah.get("jumlah_sync"),
                    sekolah.get("pd"),
                    sekolah.get("rombel"),
                    sekolah.get("ptk"),
                    sekolah.get("pegawai"),
                    sekolah.get("jml_rk"),
                    sekolah.get("jml_lab"),
                    sekolah.get("jml_perpus")
                ]
                ws.append(row)
                print(f"Data sekolah {sekolah.get('nama')} dari {nama_kecamatan}, {nama_kabupaten} ditambahkan.")
        else:
            print(f"Tidak ada data sekolah untuk kecamatan {nama_kecamatan}, {nama_kabupaten}")

# Simpan workbook ke dalam file Excel
wb.save("data_sekolah_aceh.xlsx")
print("Data berhasil disimpan dalam file Excel 'data_sekolah_aceh.xlsx'.")
