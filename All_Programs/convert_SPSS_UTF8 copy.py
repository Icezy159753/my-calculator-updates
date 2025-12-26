import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import savReaderWriter
import os
import threading

class SpssConverterApp(ttk.Window):
    def __init__(self):
        # --- 1. ตั้งค่าหน้าต่างหลัก ---
        super().__init__(themename="litera", title="โปรแกรมแปลงไฟล์ SPSS (Modern UI)")
        # กำหนดขนาดเริ่มต้น แต่ยังไม่กำหนดตำแหน่ง
        self.geometry("600x280")
        
        # --- 2. สร้างวิดเจ็ตต่างๆ ---
        self.create_widgets()

        # --- 3. จัดหน้าต่างให้อยู่กลางจอ (วิธี Manual ที่แน่นอนกว่า) ---
        self._center_window()

    def _center_window(self):
        """
        ฟังก์ชันสำหรับจัดหน้าต่างให้อยู่กึ่งกลางจอ
        เป็นวิธี manual ที่ทำงานได้กับ tkinter และ ttkbootstrap ทุกเวอร์ชัน
        """
        self.update_idletasks()  # อัปเดตเพื่อให้ได้ขนาดหน้าต่างที่แท้จริง
        
        # ดึงขนาดของหน้าต่างโปรแกรมและขนาดของหน้าจอคอมพิวเตอร์
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # คำนวณหาตำแหน่งแกน X และ Y ที่จะทำให้หน้าต่างอยู่กึ่งกลาง
        x_coordinate = int((screen_width / 2) - (window_width / 2))
        y_coordinate = int((screen_height / 2) - (window_height / 2))
        
        # ตั้งค่าตำแหน่งหน้าต่าง
        self.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    def create_widgets(self):
        # ใช้ Frame หลักเพื่อการจัดวางที่ง่ายขึ้น
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=BOTH, expand=YES)
        
        # --- แถวที่ 1: ป้ายและช่องเลือกไฟล์ ---
        ttk.Label(main_frame, text="เลือกไฟล์ SPSS ที่ต้องการซ่อม:", font=("Helvetica", 11)).grid(row=0, column=0, columnspan=2, sticky=W, pady=(0, 5))
        
        self.filepath_var = tk.StringVar()
        entry_filepath = ttk.Entry(main_frame, textvariable=self.filepath_var, font=("Helvetica", 10))
        entry_filepath.grid(row=1, column=0, sticky=EW, padx=(0, 10))
        
        self.browse_button = ttk.Button(main_frame, text="เลือกไฟล์...", command=self.select_file, bootstyle="primary")
        self.browse_button.grid(row=1, column=1, sticky=EW)

        # --- แถวที่ 2: ปุ่มแปลงไฟล์หลัก ---
        self.convert_button = ttk.Button(main_frame, text="สร้างไฟล์ SAV ใหม่ (คงค่าทั้งหมด & แก้ไข UTF-8)", command=self.start_conversion_thread, bootstyle="success")
        self.convert_button.grid(row=2, column=0, columnspan=2, pady=20, ipady=10, sticky=EW)
        
        # --- แถวที่ 3: แถบความคืบหน้า (Progress Bar) ---
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=3, column=0, columnspan=2, sticky=EW, pady=5)
        
        # --- แถวที่ 4: แถบสถานะ (Status Bar) ---
        self.status_var = tk.StringVar(value="พร้อมใช้งาน")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, font=("Helvetica", 9), bootstyle="secondary")
        status_label.grid(row=4, column=0, columnspan=2, sticky=W, pady=(10, 0))

        # กำหนดให้คอลัมน์ที่ 0 ขยายตามขนาดหน้าต่าง
        main_frame.columnconfigure(0, weight=1)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ SPSS ต้นฉบับ",
            filetypes=(("SPSS Files", "*.sav"), ("All files", "*.*"))
        )
        if file_path:
            self.filepath_var.set(file_path)
            self.status_var.set(f"ไฟล์ที่เลือก: {os.path.basename(file_path)}")

    def start_conversion_thread(self):
        self.convert_button.config(state=DISABLED)
        self.browse_button.config(state=DISABLED)
        self.progress_bar.start()
        conversion_thread = threading.Thread(target=self.clone_and_fix_spss_file)
        conversion_thread.start()

    def clone_and_fix_spss_file(self):
        spss_file_path = self.filepath_var.get()

        if not spss_file_path or not os.path.exists(spss_file_path):
            messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกไฟล์ SPSS ที่ถูกต้องก่อน")
            self.reset_ui()
            return

        try:
            self.status_var.set("กำลังอ่านไฟล์ต้นฉบับและ Metadata...")
            records = []
            metadata = {}
            with savReaderWriter.SavReader(spss_file_path, ioUtf8=True) as reader:
                metadata['varNames'] = reader.header
                metadata['varTypes'] = getattr(reader, 'varTypes', {})
                metadata['varLabels'] = getattr(reader, 'varLabels', {})
                metadata['valueLabels'] = getattr(reader, 'valueLabels', {})
                metadata['missingValues'] = getattr(reader, 'missingValues', {})
                metadata['measureLevels'] = getattr(reader, 'measureLevels', {})
                metadata['columnWidths'] = getattr(reader, 'columnWidths', {})
                metadata['formats'] = getattr(reader, 'formats', {})
                for record in reader:
                    records.append(record)
            
            self.status_var.set("กำลังเขียนไฟล์ใหม่ที่สมบูรณ์...")
            base_name = os.path.basename(spss_file_path)
            file_name_without_ext = os.path.splitext(base_name)[0]
            output_dir = os.path.dirname(spss_file_path)
            output_sav_path = os.path.join(output_dir, f"{file_name_without_ext}_modern_utf8.sav")

            with savReaderWriter.SavWriter(output_sav_path, ioUtf8=True, **metadata) as writer:
                for record in records:
                    writer.writerow(record)
            
            self.status_var.set("แปลงไฟล์สำเร็จ!")
            messagebox.showinfo("สำเร็จสมบูรณ์!", f"ไฟล์ SPSS ใหม่ที่สมบูรณ์ถูกสร้างขึ้นแล้ว\n\nบันทึกที่: {output_sav_path}")

        except Exception as e:
            self.status_var.set("เกิดข้อผิดพลาดในการแปลงไฟล์!")
            messagebox.showerror("เกิดข้อผิดพลาด", f"ประเภท: {type(e).__name__}\nข้อความ: {e}")
        finally:
            self.reset_ui()
            
    def reset_ui(self):
        self.progress_bar.stop()
        self.convert_button.config(state=NORMAL)
        self.browse_button.config(state=NORMAL)


    
# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
        # --- โค้ดที่ย้ายมาจาก if __name__ == "__main__": เดิมจะมาอยู่ที่นี่ ---
        # --- ส่วนที่ใช้รันโปรแกรม ---
    #if __name__ == "__main__":
        app = SpssConverterApp()
        app.mainloop()
        print(f"--- QUOTA_SAMPLER_INFO: QuotaSamplerApp mainloop finished. ---")

    except Exception as e:
        # ดักจับ Error ที่อาจเกิดขึ้นระหว่างการสร้างหรือรัน App
        print(f"QUOTA_SAMPLER_ERROR: An error occurred during QuotaSamplerApp execution: {e}")
        # แสดง Popup ถ้ามีปัญหา
        if 'root' not in locals() or not root.winfo_exists(): # สร้าง root ชั่วคราวถ้ายังไม่มี
            root_temp = tk.Tk()
            root_temp.withdraw()
            messagebox.showerror("Application Error (Quota Sampler)",
                               f"An unexpected error occurred:\n{e}", parent=root_temp)
            root_temp.destroy()
        else:
            messagebox.showerror("Application Error (Quota Sampler)",
                               f"An unexpected error occurred:\n{e}", parent=root) # ใช้ root ที่มีอยู่ถ้าเป็นไปได้
        sys.exit(f"Error running QuotaSamplerApp: {e}") # อาจจะ exit หรือไม่ก็ได้ ขึ้นกับการออกแบบ


# --- ส่วน Run Application เมื่อรันไฟล์นี้โดยตรง (สำหรับ Test) ---
if __name__ == "__main__":
    print("--- Running QuotaSamplerApp.py directly for testing ---")
    # (ถ้ามีการตั้งค่า DPI ด้านบน มันจะทำงานอัตโนมัติ)

    # เรียกฟังก์ชัน Entry Point ที่เราสร้างขึ้น
    run_this_app()

    print("--- Finished direct execution of QuotaSamplerApp.py ---")
# <<< END OF CHANGES >>>