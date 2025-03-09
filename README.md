# DocWizard 🧙‍♂️

DocWizard adalah sebuah script Python ajaib yang memudahkan Anda menggabungkan beberapa file DOCX menjadi satu dengan menggunakan template dan data dari file Excel. Proyek ini dirancang untuk menghasilkan dokumen secara otomatis.

## Fitur Utama

- **Penggabungan DOCX**: Menggabungkan beberapa file DOCX menjadi satu file.
- **Template Dinamis**: Menggunakan template DOCX dengan placeholder yang dapat diisi secara otomatis.
- **Data dari Excel**: Data diambil dari file Excel (`.xlsx`) untuk mengisi template.
- **Dukungan Elemen Kompleks**: Mendukung gambar, tabel, dan elemen lainnya dalam template.
- **Layout Konsisten**: Mempertahankan orientasi halaman (portrait/landscape) dan format lainnya.

## Persyaratan

- Python 3.x
- Library Python:
  - `python-docx`
  - `docxtpl`
  - `pandas`
  - `docxcompose`
  - `time`

## Instalasi

1. Clone repositori ini:
   ```bash
   https://github.com/adamusen010/DocWizard.git
# Install dependensi:
```bash
pip install -r requirements.txt
```
Cara Menggunakan
Siapkan file Template.docx dan Data.xlsx di folder proyek.
Template.docx: File template DOCX dengan placeholder (contoh: {{ nama }}, {{ Kontak }}).
Data.xlsx: File Excel yang berisi data untuk diisi ke template.
Jalankan script:
```bash
python script.py
or
py script.py
```
Hasil dokumen yang telah digabungkan akan disimpan di folder Hasil_Dokumen.
