import tkinter as tk
from tkinter import filedialog, ttk
import os

def browse_input_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
            input_label.config(text=folder_path)
            excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]
            file_count = len(excel_files)
            file_info_label.config(text=f"พบไฟล์ Excel: {file_count} ไฟล์")

def browse_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_label.config(text=folder_path)

def process_files():
    # เพิ่มโค้ดสำหรับจัดการไฟล์ Excel ได้ที่นี่
    pass

# สร้างหน้าต่างหลัก
window = tk.Tk()
window.title("Pharmacy File Manager")

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
file_info_label = ttk.Label(main_frame, text="พบไฟล์ Excel: ", justify="left")
file_info_label.pack(pady=10)

# เฟรมสำหรับปุ่มควบคุม
button_frame = ttk.Frame(main_frame)
button_frame.pack(pady=10)

process_button = ttk.Button(button_frame, text="เริ่มประมวลผล")
process_button.pack(side="left", padx=5)

window.mainloop()