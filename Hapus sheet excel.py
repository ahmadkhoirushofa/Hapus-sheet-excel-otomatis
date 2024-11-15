import os
from openpyxl import load_workbook
import shutil
import zipfile

# Fungsi untuk menghapus salah satu sheet (jika ada)
def delete_sheet_A(file_path):
    try:
        # Load workbook
        workbook = load_workbook(filename=file_path)
        
        # Cek apakah sheet A yang dimaksud ada di file
        if "Sheet A" in workbook.sheetnames:
            # Jika sheet A disembunyikan, ubah state-nya ke 'visible'
            if workbook["Sheet A"].sheet_state == 'hidden':
                workbook["Sheet A"].sheet_state = 'visible'
                
            # Hapus sheet A
            workbook.remove(workbook["Sheet A"])
            # Simpan perubahan
            workbook.save(file_path)
            print(f"Sheet A dihapus dari file {file_path}")
        else:
            print(f"Tidak ada sheet A di file {file_path}")
    except Exception as e:
        print(f"Gagal memproses file {file_path}: {e}")

# Path ke folder berisi file Excel
folder_path = "file path"
output_folder = "output path"


# Buat folder untuk menyimpan file yang sudah diproses
os.makedirs(output_folder, exist_ok=True)

# Loop untuk semua file di folder
for file_name in os.listdir(folder_path):
    if file_name.endswith(".xlsx"):
        original_file_path = os.path.join(folder_path, file_name)
        processed_file_path = os.path.join(output_folder, file_name)
        
        # Salin file ke folder output
        shutil.copy2(original_file_path, processed_file_path)
        
        # Proses file di folder output
        delete_sheet_A(processed_file_path)

# Buat file ZIP dari folder output
zip_file_path = "processed_files.zip"
with zipfile.ZipFile(zip_file_path, 'w') as zipf:
    for root, _, files in os.walk(output_folder):
        for file in files:
            file_path = os.path.join(root, file)
            zipf.write(file_path, arcname=os.path.relpath(file_path, output_folder))

print(f"File ZIP siap diunduh: {zip_file_path}")
