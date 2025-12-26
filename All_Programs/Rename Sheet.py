import customtkinter as ctk
from tkinter import filedialog
import os

class FileRenamerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("โปรแกรมเปลี่ยนชื่อไฟล์")
        self.geometry("600x550") # เพิ่มความสูงเพื่อรองรับ Textbox
        ctk.set_appearance_mode("System") # หรือ "Light", "Dark"
        ctk.set_default_color_theme("blue") # หรือ "green", "dark-blue"

        self.folder_path = ""

        # --- UI Elements ---
        # Frame หลัก
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # เลือกโฟลเดอร์
        self.folder_button = ctk.CTkButton(main_frame, text="เลือกโฟลเดอร์", command=self.select_folder)
        self.folder_button.pack(pady=(0,5))

        self.folder_label = ctk.CTkLabel(main_frame, text="ยังไม่ได้เลือกโฟลเดอร์", wraplength=500)
        self.folder_label.pack(pady=(0,10))

        # ข้อความเดิม
        ctk.CTkLabel(main_frame, text="ข้อความเดิมที่ต้องการค้นหา:").pack(pady=(10,0))
        self.old_text_entry = ctk.CTkEntry(main_frame, width=300, placeholder_text="เช่น '13 May'")
        self.old_text_entry.pack(pady=(0,10))

        # ข้อความใหม่
        ctk.CTkLabel(main_frame, text="ข้อความใหม่ที่จะใช้แทนที่:").pack()
        self.new_text_entry = ctk.CTkEntry(main_frame, width=300, placeholder_text="เช่น '14 May'")
        self.new_text_entry.pack(pady=(0,20))

        # ปุ่มเริ่มทำงาน
        self.rename_button = ctk.CTkButton(main_frame, text="เริ่มเปลี่ยนชื่อไฟล์", command=self.start_renaming)
        self.rename_button.pack(pady=10)

        # พื้นที่แสดงผลลัพธ์ (Log)
        self.log_textbox = ctk.CTkTextbox(main_frame, height=200, width=500, state="disabled")
        self.log_textbox.pack(pady=10, fill="x", expand=True)


    def select_folder(self):
        folder_selected = filedialog.askdirectory(title="เลือกโฟลเดอร์ที่ต้องการ")
        if folder_selected:
            self.folder_path = folder_selected
            self.folder_label.configure(text=f"โฟลเดอร์ที่เลือก: {self.folder_path}")
            self.log_textbox.configure(state="normal")
            self.log_textbox.delete("1.0", "end")
            self.log_textbox.insert("end", f"เลือกโฟลเดอร์: {self.folder_path}\n")
            self.log_textbox.configure(state="disabled")
        else:
            self.folder_path = ""
            self.folder_label.configure(text="ยังไม่ได้เลือกโฟลเดอร์")

    def log_message(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end") # Auto-scroll to the bottom
        self.log_textbox.configure(state="disabled")

    def start_renaming(self):
        if not self.folder_path:
            self.log_message("ข้อผิดพลาด: กรุณาเลือกโฟลเดอร์ก่อน")
            ctk.CTkMessagebox(title="ข้อผิดพลาด", message="กรุณาเลือกโฟลเดอร์ก่อน", icon="warning")
            return

        old_text = self.old_text_entry.get()
        new_text = self.new_text_entry.get()

        if not old_text:
            self.log_message("ข้อผิดพลาด: กรุณาป้อนข้อความเดิมที่ต้องการค้นหา")
            ctk.CTkMessagebox(title="ข้อผิดพลาด", message="กรุณาป้อนข้อความเดิมที่ต้องการค้นหา", icon="warning")
            return

        self.log_message(f"\nกำลังเริ่มกระบวนการเปลี่ยนชื่อ...")
        self.log_message(f"ค้นหา: '{old_text}' แทนที่ด้วย: '{new_text}'")

        renamed_count = 0
        skipped_count = 0
        error_count = 0

        try:
            for filename in os.listdir(self.folder_path):
                file_path_old = os.path.join(self.folder_path, filename)

                if os.path.isfile(file_path_old):
                    if old_text in filename:
                        new_filename = filename.replace(old_text, new_text)
                        file_path_new = os.path.join(self.folder_path, new_filename)

                        # ป้องกันการเขียนทับไฟล์ที่มีชื่อเดียวกันอยู่แล้ว (ถ้าชื่อใหม่เหมือนชื่อเก่า ก็ไม่ต้องทำอะไร)
                        if file_path_old == file_path_new:
                            self.log_message(f"ข้าม (ชื่อเหมือนเดิม): '{filename}'")
                            skipped_count +=1
                            continue

                        # ตรวจสอบว่าชื่อไฟล์ใหม่ซ้ำกับไฟล์อื่นหรือไม่
                        if os.path.exists(file_path_new):
                            self.log_message(f"ข้อผิดพลาด: ไฟล์ชื่อ '{new_filename}' มีอยู่แล้ว ไม่สามารถเปลี่ยนชื่อ '{filename}' ได้")
                            error_count +=1
                            continue

                        try:
                            os.rename(file_path_old, file_path_new)
                            self.log_message(f"เปลี่ยนชื่อ: '{filename}' -> '{new_filename}'")
                            renamed_count += 1
                        except Exception as e_rename:
                            self.log_message(f"ข้อผิดพลาดในการเปลี่ยนชื่อ '{filename}': {e_rename}")
                            error_count += 1
                    else:
                        # self.log_message(f"ข้าม (ไม่พบข้อความ): '{filename}'") # อาจจะ log เยอะไปถ้าไฟล์เยอะ
                        skipped_count += 1
                # else:
                    # self.log_message(f"ข้าม (เป็นไดเรกทอรี): '{filename}'") # ไม่จำเป็นต้อง log ทุกไดเรกทอรี
            
            summary_message = f"\n--- สรุปผล ---\n" \
                              f"เปลี่ยนชื่อสำเร็จ: {renamed_count} ไฟล์\n" \
                              f"ข้ามไป (ไม่พบข้อความ/ชื่อเดิม): {skipped_count} ไฟล์\n" \
                              f"เกิดข้อผิดพลาด: {error_count} ไฟล์"
            self.log_message(summary_message)
            if error_count > 0:
                 ctk.CTkMessagebox(title="ดำเนินการเสร็จสิ้น (มีข้อผิดพลาด)", message=summary_message, icon="warning")
            else:
                 ctk.CTkMessagebox(title="ดำเนินการเสร็จสิ้น", message=summary_message, icon="check")


        except FileNotFoundError:
            self.log_message(f"ข้อผิดพลาด: ไม่พบโฟลเดอร์ '{self.folder_path}'")
            ctk.CTkMessagebox(title="ข้อผิดพลาด", message=f"ไม่พบโฟลเดอร์ '{self.folder_path}'", icon="cancel")
        except Exception as e:
            self.log_message(f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")
            ctk.CTkMessagebox(title="ข้อผิดพลาดรุนแรง", message=f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}", icon="cancel")


# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
        # --- โค้ดที่ย้ายมาจาก if __name__ == "__main__": เดิมจะมาอยู่ที่นี่ ---

    #if __name__ == "__main__":
        app = FileRenamerApp()
        app.mainloop()       

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



