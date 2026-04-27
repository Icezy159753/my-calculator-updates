import tkinter as tk
from tkinter import filedialog, messagebox
import os

def select_file():
    file_path = filedialog.askopenfilename(
        title="เลือกไฟล์ Syntax SPSS",
        filetypes=(("SPSS Syntax Files", "*.sps"), ("All Files", "*.*"))
    )
    if file_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, file_path)

def process_file():
    input_path = entry_path.get()
    
    if not input_path or not os.path.exists(input_path):
        messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกไฟล์ .sps ที่มีอยู่จริง")
        return

    dir_name = os.path.dirname(input_path)
    base_name = os.path.basename(input_path)
    name, ext = os.path.splitext(base_name)
    output_path = os.path.join(dir_name, f"{name}_fixed{ext}")

    try:
        with open(input_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        fixed_lines = []
        for line in lines:
            # โฟกัสเฉพาะบรรทัดที่เป็น VARIABLE LABELS
            if line.strip().upper().startswith("VARIABLE LABELS"):
                # หาตำแหน่งของ " ตัวแรก และตัวสุดท้าย
                first_q = line.find('"')
                last_q = line.rfind('"')
                
                if first_q != -1 and last_q != -1 and first_q != last_q:
                    # ดึงข้อความที่อยู่ "ข้างใน" เครื่องหมายคำพูดออกมา
                    inside_text = line[first_q+1 : last_q]
                    
                    # เปลี่ยนเครื่องหมายคำพูดทุกแบบที่อยู่ข้างใน ให้เป็น ' (Single Quote)
                    fixed_inside = inside_text.replace('"', "'")
                    
                    # ประกอบร่างใหม่: เอา " หุ้มหัวท้ายเหมือนเดิม
                    fixed_line = line[:first_q] + '"' + fixed_inside + '"' + line[last_q+1:]
                    fixed_lines.append(fixed_line)
                else:
                    fixed_lines.append(line)
            else:
                fixed_lines.append(line)

        with open(output_path, 'w', encoding='utf-8') as file:
            file.writelines(fixed_lines)

        messagebox.showinfo("สำเร็จ!", f"แก้ไข Format สำเร็จ!\nได้ไฟล์รูปแบบ \"... 'brand' ...\"\n\nบันทึกไว้ที่:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"เกิดข้อผิดพลาด:\n{str(e)}")

# --- GUI Setup ---
root = tk.Tk()
root.title("SPSS Label Fixer (All DP Tool)")
root.geometry("500x150")
root.resizable(False, False)

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(fill=tk.BOTH, expand=True)

lbl_instruct = tk.Label(frame, text="เลือกไฟล์ Syntax (.sps) เพื่อแก้ปัญหา Double Quotes ซ้อนกัน:")
lbl_instruct.pack(anchor=tk.W, pady=(0, 5))

frame_input = tk.Frame(frame)
frame_input.pack(fill=tk.X, pady=(0, 15))

entry_path = tk.Entry(frame_input, width=50)
entry_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

btn_browse = tk.Button(frame_input, text="Browse...", command=select_file)
btn_browse.pack(side=tk.RIGHT)

btn_run = tk.Button(frame, text="Fix Syntax File", command=process_file, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
btn_run.pack(fill=tk.X)

root.mainloop()