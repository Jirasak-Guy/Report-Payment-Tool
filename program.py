import tkinter as tk
from tkinter import filedialog, ttk
import os
import pandas as pd

# ตัวแปร global สำหรับเก็บ path
input_path = ""
output_path = ""

def browse_input_folder():
    global input_path
    folder_path = filedialog.askdirectory()
    if folder_path:
        input_path = folder_path
        input_label.config(text=input_path)
        excel_files = [f for f in os.listdir(input_path) if f.endswith(('.xlsx', '.xls'))]
        file_info_label.config(text=f"พบไฟล์ Excel: {len(excel_files)} ไฟล์")

def browse_output_folder():
    global output_path
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_path = folder_path
        output_label.config(text=output_path)

def process_files():
    if not input_path or not output_path:
        file_info_label.config(text="กรุณาเลือกโฟลเดอร์ทั้งสอง")
        return

    excel_files = [f for f in os.listdir(input_path) if f.endswith(('.xlsx', '.xls'))]
    if not excel_files:
        file_info_label.config(text="ไม่พบไฟล์ Excel ในโฟลเดอร์")
        return
    
    # สร้างโฟลเดอร์ใหม่ใน output path
    input_folder_name = os.path.basename(input_path)
    new_output_folder = os.path.join(output_path, f"{input_folder_name}_แก้ไขแล้ว")
    os.makedirs(new_output_folder, exist_ok=True)

    progress_bar["maximum"] = len(excel_files)
    progress_bar["value"] = 0

    for idx, file_name in enumerate(excel_files):
        file_path = os.path.join(input_path, file_name)
        try:
            df = pd.read_excel(file_path, dtype={'ราคาต่อหน่วย': 'float64', 'จำนวน': 'float64', 'ราคาสุทธิ': 'float64'})

            expected_columns = ['ลำดับ', 'รายการ', 'วันที่', 'ราคาต่อหน่วย', 'จำนวน', 'ราคาสุทธิ']
            if not all(col in df.columns for col in expected_columns):
                file_info_label.config(text=f"ไฟล์ {file_name} ไม่มีคอลัมน์ที่ต้องการ")
                continue

            # สร้างลำดับใหม่สำหรับบิล (ORR)
            new_order = 1
            for i, row in df.iterrows():
                if isinstance(row['รายการ'], str) and row['รายการ'].startswith('ORR'):
                    df.at[i, 'ลำดับ'] = new_order
                    new_order += 1
                else:
                    df.at[i, 'ลำดับ'] = ''

            # ตั้งค่าการแสดงผลใน Excel ให้ไม่มีทศนิยม แต่ค่าเดิมยังอยู่
            output_file_name = f"{os.path.splitext(file_name)[0]}_แก้ไขแล้ว.xlsx"
            output_file = os.path.join(new_output_folder, output_file_name)

            with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")
                workbook = writer.book
                worksheet = writer.sheets["Sheet1"]

                # กำหนดรูปแบบตัวเลขให้ไม่มีทศนิยม
                format_no_decimal = workbook.add_format({"num_format": "0"})
                
                # หา index ของคอลัมน์ที่ต้องการกำหนดค่า
                column_indexes = {col_name: idx for idx, col_name in enumerate(df.columns)}

                # ใช้ format_no_decimal กับคอลัมน์ที่เกี่ยวข้อง
                for col in ['ราคาต่อหน่วย', 'จำนวน', 'ราคาสุทธิ']:
                    if col in column_indexes:
                        col_letter = chr(65 + column_indexes[col])  # แปลง index เป็นตัวอักษร (A, B, C, ...)
                        worksheet.set_column(f"{col_letter}:{col_letter}", None, format_no_decimal)

            progress_bar["value"] += 1
            window.update_idletasks()

        except Exception as e:
            file_info_label.config(text=f"เกิดข้อผิดพลาดกับ {file_name}: {str(e)}")
            continue

    file_info_label.config(text="ประมวลผลเสร็จสิ้น!")
    progress_bar["value"] = len(excel_files)

# สร้างหน้าต่างหลัก
window = tk.Tk()
window.title("Pharmacy File Manager")

window_width, window_height = 450, 250
screen_width, screen_height = window.winfo_screenwidth(), window.winfo_screenheight()
x_coordinate, y_coordinate = (screen_width - window_width) // 2, (screen_height - window_height) // 2
window.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

main_frame = ttk.Frame(window, padding="10")
main_frame.pack(fill="both", expand=True)

# Input Folder
ttk.Label(main_frame, text="Input Folder (โฟลเดอร์ที่มีไฟล์ Excel)").pack(pady=5)
input_frame = ttk.Frame(main_frame)
input_frame.pack(fill="x")
input_label = ttk.Label(input_frame, text="ยังไม่ได้เลือกโฟลเดอร์", wraplength=350)
input_label.pack(side="left", padx=5, fill="x", expand=True)
ttk.Button(input_frame, text="Browse...", command=browse_input_folder).pack(side="right")

# Output Folder
ttk.Label(main_frame, text="Output Folder (โฟลเดอร์ปลายทาง)").pack(pady=5)
output_frame = ttk.Frame(main_frame)
output_frame.pack(fill="x")
output_label = ttk.Label(output_frame, text="ยังไม่ได้เลือกโฟลเดอร์", wraplength=350)
output_label.pack(side="left", padx=5, fill="x", expand=True)
ttk.Button(output_frame, text="Browse...", command=browse_output_folder).pack(side="right")

# ข้อมูลเพิ่มเติม
file_info_label = ttk.Label(main_frame, text="กรุณาเลือกโฟลเดอร์", justify="left")
file_info_label.pack(pady=10)

# Progress Bar
progress_bar = ttk.Progressbar(main_frame, length=300, mode="determinate")
progress_bar.pack(pady=5)

# ปุ่มเริ่มประมวลผล
button_frame = ttk.Frame(main_frame)
button_frame.pack(pady=10)
ttk.Button(button_frame, text="เริ่มประมวลผล", command=process_files).pack(side="left", padx=5)

window.mainloop()