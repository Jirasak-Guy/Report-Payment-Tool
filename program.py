import tkinter as tk
from tkinter import filedialog, ttk
import os
import pandas as pd
from openpyxl import load_workbook

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
    # ตรวจสอบว่าเลือกโฟลเดอร์ครบหรือไม่
    if not input_path or not output_path:
        file_info_label.config(text="กรุณาเลือกโฟลเดอร์ทั้งสอง")
        return
    # ดึงรายการไฟล์ Excel จาก input folder
    excel_files = [f for f in os.listdir(input_path) if f.endswith(('.xlsx', '.xls'))]
    if not excel_files:
        file_info_label.config(text="ไม่พบไฟล์ Excel ในโฟลเดอร์")
        return
    
    # สร้างโฟลเดอร์ใหม่ใน output path โดยใช้ชื่อ input folder + "แก้ไขแล้ว"
    input_folder_name = os.path.basename(input_path)
    new_output_folder = os.path.join(output_path, f"{input_folder_name}_แก้ไขแล้ว")
    os.makedirs(new_output_folder, exist_ok=True)

    for file_name in excel_files:
        file_path = os.path.join(input_path, file_name)
        try:
            # อ่านไฟล์ Excel
            df = pd.read_excel(file_path)
            
            # ตรวจสอบว่ามีคอลัมน์ครบตามที่ระบุ
            expected_columns = ['ลำดับ', 'รายการ', 'วันที่', 'ราคาต่อหน่วย', 'จำนวน', 'ราคาสุทธิ']
            if not all(col in df.columns for col in expected_columns):
                file_info_label.config(text=f"ไฟล์ {file_name} ไม่มีคอลัมน์ที่ต้องการ")
                continue
            
            # สร้างลำดับใหม่สำหรับบิล (ORR)
            new_order = 1
            for idx, row in df.iterrows():
                if isinstance(row['รายการ'], str) and row['รายการ'].startswith('ORR'):
                    df.at[idx, 'ลำดับ'] = new_order
                    new_order += 1
                else:
                    # ถ้าไม่ใช่บิล (เช่น เป็นสินค้า P-xxx) ให้ลำดับเป็นว่าง
                    df.at[idx, 'ลำดับ'] = ''
                    
            
            # บันทึกไฟล์ใหม่ในโฟลเดอร์ที่สร้าง โดยใช้ชื่อเดิม + "แก้ไขแล้ว"
            output_file_name = f"{os.path.splitext(file_name)[0]}_แก้ไขแล้ว{os.path.splitext(file_name)[1]}"
            output_file = os.path.join(new_output_folder, output_file_name)
            df.to_excel(output_file, index=False)

            # เปิดไฟล์ Excel ด้วย openpyxl เพื่อกำหนดรูปแบบและความกว้าง
            workbook = load_workbook(output_file)
            worksheet = workbook.active

            # กำหนดรูปแบบตัวเลขในคอลัมน์ 'ราคาต่อหน่วย', 'จำนวน', 'ราคาสุทธิ'
            numeric_columns = ['ราคาต่อหน่วย', 'จำนวน', 'ราคาสุทธิ']
            col_indices = [df.columns.get_loc(col) + 1 for col in numeric_columns]  # +1 เพราะ openpyxl เริ่มที่ 1
            for col_idx in col_indices:
                for row in range(2, worksheet.max_row + 1):  # เริ่มที่แถว 2 (ข้าม header)
                    cell = worksheet.cell(row=row, column=col_idx)
                    # ถ้าค่าในเซลล์เป็น 0 (หรือใกล้เคียง) ให้กำหนดรูปแบบพิเศษ
                    if cell.value is not None and isinstance(cell.value, (int, float)) and abs(cell.value) < 0.0001:
                        cell.number_format = '"-"'
                    else:
                        cell.number_format = '0'  # รูปแบบปกติ (จำนวนเต็ม)

            # คำนวณความกว้างอัตโนมัติสำหรับแต่ละคอลัมน์ (ไม่คูณ 1.2)
            for i, column in enumerate(df.columns, 1):
                col_letter = chr(64 + i)  # แปลงเลขคอลัมน์เป็นตัวอักษร (A, B, C, ...)
                # หาความยาวสูงสุดของข้อความในคอลัมน์ (รวม header)
                max_length = max(df[column].astype(str).apply(len).max(), len(str(column)))
                # ถ้าคอลัมน์อยู่ใน numeric_columns และมีค่า 0 ให้เผื่อความกว้างสำหรับ "-"
                if column in numeric_columns:
                    max_length = max(max_length, 1)  # ความกว้างอย่างน้อย 1 สำหรับ "-"
                worksheet.column_dimensions[col_letter].width = max_length

            # บันทึกไฟล์ที่ปรับรูปแบบและความกว้างแล้ว
            workbook.save(output_file)
            
            file_info_label.config(text=f"ประมวลผลสำเร็จ: {file_name}")
        
        except Exception as e:
            file_info_label.config(text=f"เกิดข้อผิดพลาดกับ {file_name}: {str(e)}")
            continue

# สร้างหน้าต่างหลัก
window = tk.Tk()
window.title("Report Payment Tool")

# กำหนดขนาดหน้าต่าง
window_width = 400
window_height = 220

# คำนวณตำแหน่งให้อยู่กลางจอ
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x_coordinate = int((screen_width / 2) - (window_width / 2))
y_coordinate = int((screen_height / 2) - (window_height / 2))

# ตั้งค่าขนาดและตำแหน่ง
window.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
window.minsize(400, 220)

# เฟรมหลัก
main_frame = ttk.Frame(window, padding="10")
main_frame.pack(fill="both", expand=True)

# เลเบลและปุ่มสำหรับ Input Folder
ttk.Label(main_frame, text="Input Folder (โฟลเดอร์ที่มีไฟล์ Excel)").pack(pady=5)
input_frame = ttk.Frame(main_frame)
input_frame.pack(fill="x")

input_label = ttk.Label(input_frame, text="ยังไม่ได้เลือกโฟลเดอร์", wraplength=300)
input_label.pack(side="left", padx=5, fill="x", expand=True)

input_button = ttk.Button(input_frame, text="Browse...", command=browse_input_folder)
input_button.pack(side="right")

# เลเบลและปุ่มสำหรับ Output Folder
ttk.Label(main_frame, text="Output Folder (โฟลเดอร์ปลายทาง)").pack(pady=5)
output_frame = ttk.Frame(main_frame)
output_frame.pack(fill="x")

output_label = ttk.Label(output_frame, text="ยังไม่ได้เลือกโฟลเดอร์", wraplength=300)
output_label.pack(side="left", padx=5, fill="x", expand=True)

output_button = ttk.Button(output_frame, text="Browse...", command=browse_output_folder)
output_button.pack(side="right")

# ข้อมูลเพิ่มเติม
file_info_label = ttk.Label(main_frame, text="กรุณาเลือกโฟลเดอร์", justify="left")
file_info_label.pack(pady=10)

# เฟรมสำหรับปุ่มควบคุม
button_frame = ttk.Frame(main_frame)
button_frame.pack(pady=10)

process_button = ttk.Button(button_frame, text="เริ่มประมวลผล", command=process_files)
process_button.pack(side="left", padx=5)

window.mainloop()