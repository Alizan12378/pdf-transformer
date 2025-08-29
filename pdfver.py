import os
import sys
import shutil
import img2pdf
from PIL import Image
from fpdf import FPDF
from docx2pdf import convert as docx2pdf_convert
import pythoncom
from win32com import client

def word_to_pdf(input_path, output_path):
    try:
        docx2pdf_convert(input_path, output_path)
        print(f"Converted Word: {input_path}")
    except Exception as e:
        print(f"Error converting Word {input_path}: {e}")

def excel_to_pdf(input_path, output_path):
    try:
        pythoncom.CoInitialize()
        excel = client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(input_path)
        ws = wb.Worksheets(1)

        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1

        wb.ExportAsFixedFormat(0, output_path)
        wb.Close()
        excel.Quit()
        print(f"Converted Excel (fit to one page): {input_path}")
    except Exception as e:
        print(f"Error converting Excel {input_path}: {e}")

def img_to_pdf(input_path, output_path):
    try:
        with open(output_path, 'wb') as f:
            f.write(img2pdf.convert(input_path))
        print(f"Converted Image: {input_path}")
    except Exception as e:
        print(f"Error converting Image {input_path}: {e}")

def txt_to_pdf(input_path, output_path):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        with open(input_path, 'r', encoding='utf-8') as f:
            for line in f:
                pdf.cell(0, 10, txt=line.strip(), ln=True)
        pdf.output(output_path)
        print(f"Converted Text: {input_path}")
    except Exception as e:
        print(f"Error converting Text {input_path}: {e}")

def ensure_unique_path(folder, filename):
    base, ext = os.path.splitext(filename)
    count = 1
    new_path = os.path.join(folder, filename)
    while os.path.exists(new_path):
        new_filename = f"{base}_{count}{ext}"
        new_path = os.path.join(folder, new_filename)
        count += 1
    return new_path

def ensure_dir_exists(path):
    if not os.path.exists(path):
        os.makedirs(path)

def get_application_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def convert_and_copy_flat(src_folder, dest_folder):
    ensure_dir_exists(dest_folder)
    for root, dirs, files in os.walk(src_folder):
        for file in files:
            full_src_path = os.path.join(root, file)

            name, ext = os.path.splitext(file)
            ext = ext.lower()

            if ext in ['.doc', '.docx']:
                target_pdf = ensure_unique_path(dest_folder, name + ".pdf")
                word_to_pdf(full_src_path, target_pdf)
            elif ext in ['.xls', '.xlsx']:
                target_pdf = ensure_unique_path(dest_folder, name + ".pdf")
                excel_to_pdf(full_src_path, target_pdf)
            elif ext in ['.jpg', '.jpeg', '.png', '.bmp']:
                target_pdf = ensure_unique_path(dest_folder, name + ".pdf")
                img_to_pdf(full_src_path, target_pdf)
            elif ext == '.txt':
                target_pdf = ensure_unique_path(dest_folder, name + ".pdf")
                txt_to_pdf(full_src_path, target_pdf)
            elif ext == '.pdf':
                target_pdf = ensure_unique_path(dest_folder, file)
                shutil.copy2(full_src_path, target_pdf)
                print(f"Copied PDF: {full_src_path}")
            else:
                print(f"不支援格式忽略: {full_src_path}")

if __name__ == "__main__":
    folder_to_process = input("請輸入要掃描的資料夾路徑：")
    application_path = get_application_path()
    output_folder = os.path.join(application_path, "Allfile")

    ensure_dir_exists(output_folder)
    convert_and_copy_flat(folder_to_process, output_folder)
