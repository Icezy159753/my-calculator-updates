import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pyreadstat
import pyperclip
from collections import defaultdict
import re
import sys

# --- Natural Sort Helper Function ---
def natural_sort_key(s):
    """
    A key for sorting strings in a 'natural' order (e.g., 'a10' comes after 'a2').
    """
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

class SPSS_MRSET_Generator:
    def __init__(self, root):
        self.root = root
        self.root.title("SPSS MRSET Auto-Generator v5.0 (Natural Sort)")

        # --- โค้ดสำหรับจัดให้โปรแกรมอยู่กลางจอ ---
        window_width = 700
        window_height = 650

        # ดึงขนาดของหน้าจอ
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # คำนวณตำแหน่งกึ่งกลาง
        center_x = int(screen_width/2 - window_width / 2)
        center_y = int(screen_height/2 - window_height / 2)

        # ตั้งค่าขนาดและตำแหน่งของหน้าต่าง
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        # --- สิ้นสุดส่วนจัดกลางจอ ---

        self.root.resizable(False, False)

        self.file_path = None

        # --- GUI Elements ---
        # Frame for file selection
        file_frame = tk.LabelFrame(root, text="1. เลือกไฟล์ SPSS", padx=10, pady=10)
        file_frame.pack(padx=10, pady=5, fill="x")

        self.file_label = tk.Label(file_frame, text="ยังไม่ได้เลือกไฟล์", fg="grey", width=70, anchor="w")
        self.file_label.pack(side=tk.LEFT, padx=5)
        self.select_button = tk.Button(file_frame, text="เลือกไฟล์ (.sav)", command=self.select_file)
        self.select_button.pack(side=tk.RIGHT)

        # Frame for MRSET options
        options_frame = tk.LabelFrame(root, text="2. กำหนดค่า", padx=10, pady=10)
        options_frame.pack(padx=10, pady=5, fill="x")
        
        options_frame.columnconfigure(1, weight=1)

        # Set Type (MDSET/MCSET)
        tk.Label(options_frame, text="ประเภทชุดคำตอบ:").grid(row=0, column=0, sticky="w", pady=5)
        self.set_type = tk.StringVar(value="MDGROUP")
        tk.Radiobutton(options_frame, text="Multiple Dichotomy (MDSET)", variable=self.set_type, value="MDGROUP", command=self.toggle_counted_value).grid(row=1, column=0, sticky="w", padx=10)
        tk.Radiobutton(options_frame, text="Multiple Category (MCSET)", variable=self.set_type, value="MCGROUP", command=self.toggle_counted_value).grid(row=1, column=1, sticky="w")

        # MA Identifier
        tk.Label(options_frame, text="สัญลักษณ์บ่งชี้ข้อ MA (เช่น _O, $):").grid(row=2, column=0, sticky="w", pady=5)
        self.identifier_entry = tk.Entry(options_frame, width=40)
        self.identifier_entry.grid(row=2, column=1, sticky="w")

        # Counted Value (for MDGROUP)
        self.value_label = tk.Label(options_frame, text="ค่าที่ต้องการนับ (Counted Value):")
        self.value_label.grid(row=3, column=0, sticky="w", pady=5)
        self.value_entry = tk.Entry(options_frame, width=40)
        self.value_entry.insert(0, "1")
        self.value_entry.grid(row=3, column=1, sticky="w")
        
        # Use TO keyword option
        self.use_to_keyword = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, text="ใช้ 'TO' สำหรับตัวแปรที่เรียงติดกัน (กระชับ Syntax)", variable=self.use_to_keyword).grid(row=4, column=0, columnspan=2, sticky="w", pady=5, padx=5)

        # Generate Button
        self.generate_button = tk.Button(root, text="3. สร้างโค้ด Syntax ทั้งหมด", command=self.generate_syntax, font=("Helvetica", 10, "bold"))
        self.generate_button.pack(pady=10)

        # Output text area
        output_frame = tk.LabelFrame(root, text="4. โค้ด SPSS Syntax ที่ได้ (ทั้งหมด)", padx=10, pady=10)
        output_frame.pack(padx=10, pady=5, fill="both", expand=True)

        self.syntax_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, height=10, width=80)
        self.syntax_text.pack(fill="both", expand=True)
        self.syntax_text.insert(tk.END, "ผลลัพธ์จะแสดงที่นี่...\n\nฟีเจอร์ใหม่ (v5.0):\nอัปเกรดการเรียงลำดับเป็น 'Natural Sort' เพื่อให้เข้าใจตัวแปร\nเช่น s1_O10 ได้อย่างถูกต้อง")
        self.syntax_text.config(state="disabled")

        self.copy_button = tk.Button(root, text="คัดลอกโค้ดทั้งหมด", command=self.copy_to_clipboard)
        self.copy_button.pack(pady=5)

    def toggle_counted_value(self):
        if self.set_type.get() == "MCGROUP":
            self.value_label.grid_remove()
            self.value_entry.grid_remove()
        else: # MDGROUP
            self.value_label.grid()
            self.value_entry.grid()

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="เลือกไฟล์ SPSS",
            filetypes=(("SPSS Data Files", "*.sav"), ("All files", "*.*"))
        )
        if self.file_path:
            filename = self.file_path.split('/')[-1]
            self.file_label.config(text=filename, fg="black")
        else:
            self.file_label.config(text="ยังไม่ได้เลือกไฟล์", fg="grey")

    def generate_syntax(self):
        if not self.file_path:
            messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกไฟล์ SPSS ก่อน")
            return

        ma_identifier = self.identifier_entry.get().strip()
        if not ma_identifier:
            messagebox.showerror("ข้อผิดพลาด", "กรุณาระบุ 'สัญลักษณ์บ่งชี้ข้อ MA'")
            return

        set_type_val = self.set_type.get()
        use_to = self.use_to_keyword.get()

        try:
            _, meta = pyreadstat.read_sav(self.file_path, metadataonly=True)
            all_vars_ordered = meta.column_names
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาดในการอ่านไฟล์", f"ไม่สามารถอ่านไฟล์ SPSS ได้:\n{e}")
            return

        ma_vars = [var for var in all_vars_ordered if ma_identifier in var]
        if not ma_vars:
            messagebox.showwarning("ไม่พบตัวแปร", f"ไม่พบตัวแปรที่มีสัญลักษณ์ '{ma_identifier}' ในไฟล์นี้")
            return
            
        grouped_vars = defaultdict(list)
        for var in ma_vars:
            prefix = var.split(ma_identifier)[0]
            grouped_vars[prefix].append(var)

        all_syntax_blocks = []
        for prefix, var_list in grouped_vars.items():
            set_name = f"${prefix}"
            
            # --- USE NATURAL SORT KEY HERE ---
            sorted_vars = sorted(var_list, key=natural_sort_key)

            first_var = sorted_vars[0]
            try:
                var_index = all_vars_ordered.index(first_var)
                label = meta.column_labels[var_index] or prefix
            except (ValueError, IndexError):
                label = prefix
            
            label = label.replace("'", "")

            vars_string = ""
            is_consecutive = False
            if use_to and len(sorted_vars) > 1:
                try:
                    start_index = all_vars_ordered.index(sorted_vars[0])
                    end_index = all_vars_ordered.index(sorted_vars[-1])
                    if all_vars_ordered[start_index : end_index + 1] == sorted_vars:
                        is_consecutive = True
                except ValueError:
                    is_consecutive = False
            
            if is_consecutive:
                vars_string = f"{sorted_vars[0]} TO {sorted_vars[-1]}"
            else:
                vars_string = "\n    ".join(sorted_vars)

            syntax_block = ""
            if set_type_val == "MDGROUP":
                counted_value = self.value_entry.get().strip()
                if not counted_value:
                    messagebox.showerror("ข้อผิดพลาด", "กรุณาระบุ 'ค่าที่ต้องการนับ' สำหรับ MDSET")
                    return
                syntax_block = (f"MRSETS\n  /MDGROUP NAME={set_name} LABEL='{label}'\n"
                                f"    VARIABLES={vars_string}\n    VALUE={counted_value}.")
            elif set_type_val == "MCGROUP":
                syntax_block = (f"MRSETS\n  /MCGROUP NAME={set_name} LABEL='{label}'\n"
                                f"    VARIABLES={vars_string}.")

            all_syntax_blocks.append(syntax_block)
        
        final_syntax = "\n\n".join(all_syntax_blocks)
        
        self.syntax_text.config(state="normal")
        self.syntax_text.delete(1.0, tk.END)
        self.syntax_text.insert(tk.END, f"* Generated {len(grouped_vars)} Multiple Response Set(s).\n\n" + final_syntax)
        self.syntax_text.config(state="disabled")
        messagebox.showinfo("สำเร็จ", f"สร้างโค้ดสำเร็จ! พบคำถาม MA ทั้งหมด {len(grouped_vars)} กลุ่ม")

    def copy_to_clipboard(self):
        self.syntax_text.config(state="normal")
        code_to_copy = self.syntax_text.get(1.0, tk.END).strip()
        if code_to_copy and "ผลลัพธ์จะแสดงที่นี่" not in code_to_copy:
            pyperclip.copy(code_to_copy)
            messagebox.showinfo("คัดลอกแล้ว", "คัดลอกโค้ดลงในคลิปบอร์ดเรียบร้อยแล้ว")
        else:
            messagebox.showwarning("ไม่มีโค้ด", "ยังไม่มีโค้ดให้คัดลอก กรุณากดสร้างโค้ดก่อน")
        self.syntax_text.config(state="disabled")



# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
    # --- ส่วนที่ใช้รันโปรแกรม ---
    #if __name__ == "__main__":
        root = tk.Tk()
        app = SPSS_MRSET_Generator(root)
        root.mainloop()

        
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
