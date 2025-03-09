import pandas as pd
import os
from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer
import time  

# 1. Tentukan folder output
output_folder = "Hasil_Dokumen"  
os.makedirs(output_folder, exist_ok=True)

# 2. Baca file Excel
excel_file = "Data.xlsx"  
df = pd.read_excel(excel_file)
df = df.sort_index()

# 3. Path template Word
template_path = "Template.docx"
start_time = time.time()  
# 4. Buat dokumen gabungan menggunakan Composer
combined_doc = Document()  
composer = Composer(combined_doc)  

for index, row in df.iterrows():
    
    doc = DocxTemplate(template_path)

    # Ambil data dari Excel
    data = {
        "nama": row["nama"],
        "Kontak": str(row["Kontak"]),
        "alamat_pengiriman_serdik": str(row["alamat_pengiriman_serdik"]),
        "Kode_Pos": str(row["Kode_Pos"]),
       # "Desa": str(row["Desa"]),
        "Kecamatan": row["Kecamatan"],
        "Kabupaten": row["Kabupaten"],
        "Provinsi": row["Provinsi"],
        "No_Blangko": str(row["No_Blangko"]),
        "Ukuran_Baju": row["Ukuran_Baju"]
    }

    doc.render(data)

    # Simpan file sementara
    temp_output = os.path.join(output_folder, f"temp_{index}.docx")
    doc.save(temp_output)
    temp_doc = Document(temp_output)
    composer.append(temp_doc)
    os.remove(temp_output)

# 5. Simpan hasil sebagai satu file Word
output_path = os.path.join(output_folder, "Output.docx")
composer.save(output_path)

end_time = time.time()  
execution_time = end_time - start_time

# Outputkan
print(f"Dokumen berhasil dibuat di folder: {output_folder}")
print(f"File hasil: {output_path}")
print(f"Waktu eksekusi: {execution_time:.2f} detik")
