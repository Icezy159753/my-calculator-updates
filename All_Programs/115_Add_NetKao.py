import tkinter as tk
from tkinter import ttk
import sv_ttk # <-- 1. Import theme ที่ติดตั้งไว้

class ModernDataEntryApp:
    def __init__(self, root):
        """
        Constructor for the ModernDataEntryApp class.
        Initializes the main window with a modern theme and creates all the UI widgets.
        """
        self.root = root
        self.root.title("โปรแกรมกรอกข้อมูล (Modern UI)")
        self.root.geometry("900x850") # ขยายหน้าต่างให้กว้างขึ้นเล็กน้อย
        self.root.minsize(800, 750) # กำหนดขนาดต่ำสุดเพื่อไม่ให้ layout พัง

        # --- 2. กำหนด Style และ Font ที่จะใช้ทั้งโปรแกรม ---
        self.header_font = ("Calibri", 14, "bold")
        self.label_font = ("Calibri", 11)
        self.entry_font = ("Calibri", 11)

        # --- สร้าง Frame หลักสำหรับบรรจุทุกอย่าง ---
        # เพิ่ม padding ด้านนอกเพื่อให้มีขอบสวยงาม
        main_container = ttk.Frame(self.root, padding="20")
        main_container.pack(fill=tk.BOTH, expand=True)

        # --- เพิ่มหัวข้อหลักของโปรแกรม ---
        app_title = ttk.Label(
            main_container,
            text="ระบบบันทึกข้อมูลการประเมินผล",
            font=("Calibri", 22, "bold"),
            anchor="center"
        )
        app_title.pack(fill=tk.X, pady=(0, 20))

        # --- สร้าง 3 ส่วนหลักของโปรแกรม ---
        self.create_top_section(main_container)
        self.create_code_frame_section(main_container)
        self.create_ce_oe_scale_section(main_container)
        self.create_action_buttons(main_container)

    def create_top_section(self, parent_container):
        """สร้างตารางสำหรับ Net และ Code Net ในดีไซน์ที่ดูง่ายขึ้น"""
        # ใช้ LabelFrame เพื่อจัดกลุ่ม แต่ปรับ font ให้ดูเด่นขึ้น
        top_frame = ttk.LabelFrame(
            parent_container,
            text=" Net / Code Net ", # เพิ่ม space รอบข้อความ
            padding="15",
            labelwidget=ttk.Label(text=" Net / Code Net ", font=self.header_font)
        )
        top_frame.pack(fill=tk.X, expand=True, pady=10)

        top_frame.columnconfigure(1, weight=1) # ให้คอลัมน์ "Net" ขยายได้

        # --- หัวข้อตาราง ---
        headers = ["No.", "Net Description", "Code Net"]
        for i, header in enumerate(headers):
            lbl = ttk.Label(top_frame, text=header, font=self.label_font + ("bold",))
            lbl.grid(row=0, column=i, padx=5, pady=(5, 10), sticky='w')

        # เพิ่มเส้นคั่นใต้หัวข้อ
        separator = ttk.Separator(top_frame, orient='horizontal')
        separator.grid(row=1, column=0, columnspan=3, sticky='ew', pady=(0, 10))


        # --- ข้อมูลในตาราง ---
        net_data = [
            "+++++ POSITIVE +++++(ด้านบวก)",
            "++++ Some positive points ++++ (ตอบรสชาติด้านบวก 1,2 แต่ให้เหตุผลด้านบวก)",
            "*** The different good performance ***(ตอบ Code 3 และไม่ตอบ Code ส่วนที่บอกว่าทั้งคู่เหมือนกัน)",
            "----- NEGATIVE - - - - -(ด้านลบ)",
            "----- Some negative points - - - -(ตอบรสชาติด้านบวก 5,4 แต่ให้เหตุผลด้านลบ)",
            "*** The same good performance ***(ตอบ Codeเฉพาะส่วนที่เป็นทั้งคู่เหมือนกัน)"
        ]

        # --- สร้างแถวข้อมูลและช่องกรอก ---
        for i in range(6):
            # No.
            no_lbl = ttk.Label(top_frame, text=f"{i + 1}.", font=self.label_font)
            no_lbl.grid(row=i + 2, column=0, padx=5, pady=5, sticky='n')

            # Net Description
            net_lbl = ttk.Label(top_frame, text=net_data[i], wraplength=450, justify=tk.LEFT, font=self.label_font)
            net_lbl.grid(row=i + 2, column=1, padx=5, pady=5, sticky='w')

            # Code Net Entry
            code_net_entry = ttk.Entry(top_frame, width=20, font=self.entry_font)
            code_net_entry.grid(row=i + 2, column=2, padx=5, pady=5, sticky='w')

    def create_code_frame_section(self, parent_container):
        """สร้างส่วนของ Code Frame ในดีไซน์ที่ดูง่ายขึ้น"""
        code_frame = ttk.LabelFrame(
            parent_container,
            text=" Code Frame ",
            padding="15",
            labelwidget=ttk.Label(text=" Code Frame ", font=self.header_font)
        )
        code_frame.pack(fill=tk.X, expand=True, pady=10)

        code_frame.columnconfigure(0, weight=1) # ให้ Label ขยายได้

        code_frame_labels = [
            "Code ด้านบวกทั้งหมด",
            "Code ด้านลบทั้งหมด",
            "Code ด้านคู่เหมือนกัน"
        ]
        for i, text in enumerate(code_frame_labels):
            lbl = ttk.Label(code_frame, text=text, font=self.label_font)
            lbl.grid(row=i, column=0, padx=5, pady=5, sticky='w')

            entry = ttk.Entry(code_frame, width=20, font=self.entry_font)
            entry.grid(row=i, column=1, padx=5, pady=5, sticky='e')

    def create_ce_oe_scale_section(self, parent_container):
        """สร้างตารางสำหรับ CE, OE, Scale ในดีไซน์ที่ดูง่ายขึ้น"""
        ce_oe_scale_frame = ttk.LabelFrame(
            parent_container,
            text=" CE / OE / Scale ",
            padding="15",
            labelwidget=ttk.Label(text=" CE / OE / Scale ", font=self.header_font)
        )
        ce_oe_scale_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # กำหนดให้ทั้ง 3 คอลัมน์ขยายเท่าๆ กัน
        ce_oe_scale_frame.columnconfigure(0, weight=1)
        ce_oe_scale_frame.columnconfigure(1, weight=1)
        ce_oe_scale_frame.columnconfigure(2, weight=1)

        # --- สร้าง Dropdown (Combobox) พร้อม Label ---
        category_label = ttk.Label(ce_oe_scale_frame, text="เลือกประเภท:", font=self.label_font)
        category_label.grid(row=0, column=0, padx=5, pady=(0, 15), sticky='w')

        dropdown_options = [
            "++++ POSITIVE ++++", "++++ Some positive points ++++",
            "*** The different good performance ***", "---- NEGATIVE ----",
            "---- Some negative points ----", "*** The same good performance ***"
        ]
        category_combo = ttk.Combobox(
            ce_oe_scale_frame, values=dropdown_options,
            state="readonly", font=self.entry_font, width=35
        )
        category_combo.grid(row=1, column=0, columnspan=2, padx=5, pady=(0, 15), sticky='w')
        category_combo.set("เลือกหมวดหมู่...")

        # --- หัวข้อตาราง ---
        headers = ["CE", "OE", "Scale"]
        for i, header in enumerate(headers):
            lbl = ttk.Label(ce_oe_scale_frame, text=header, font=self.label_font + ("bold",))
            lbl.grid(row=2, column=i, padx=5, pady=5, sticky='w')

        # เพิ่มเส้นคั่นใต้หัวข้อ
        separator = ttk.Separator(ce_oe_scale_frame, orient='horizontal')
        separator.grid(row=3, column=0, columnspan=3, sticky='ew', pady=(0, 10))

        # --- สร้างช่องกรอก 10 แถว ---
        for i in range(10):
            row_index = i + 4
            ce_entry = ttk.Entry(ce_oe_scale_frame, font=self.entry_font)
            ce_entry.grid(row=row_index, column=0, padx=5, pady=3, sticky='ew')

            oe_entry = ttk.Entry(ce_oe_scale_frame, font=self.entry_font)
            oe_entry.grid(row=row_index, column=1, padx=5, pady=3, sticky='ew')

            scale_entry = ttk.Entry(ce_oe_scale_frame, font=self.entry_font)
            scale_entry.grid(row=row_index, column=2, padx=5, pady=3, sticky='ew')

    def create_action_buttons(self, parent_container):
        """สร้างปุ่มสำหรับดำเนินการ เช่น บันทึก หรือ ล้างข้อมูล"""
        button_frame = ttk.Frame(parent_container, padding=(0, 10))
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # Style สำหรับปุ่มหลัก (Accent button)
        style = ttk.Style()
        style.configure("Accent.TButton", font=self.label_font + ("bold",))

        # ปุ่ม Clear
        clear_button = ttk.Button(
            button_frame,
            text="ล้างข้อมูล",
            command=self.clear_data # สร้างฟังก์ชันนี้เพิ่ม
        )
        clear_button.pack(side=tk.RIGHT, padx=5)

        # ปุ่ม Save
        save_button = ttk.Button(
            button_frame,
            text="บันทึกข้อมูล",
            style="Accent.TButton", # ใช้ style พิเศษเพื่อให้เด่นขึ้น
            command=self.save_data # สร้างฟังก์ชันนี้เพิ่ม
        )
        save_button.pack(side=tk.RIGHT)

    def save_data(self):
        """ฟังก์ชันตัวอย่างสำหรับการบันทึกข้อมูล"""
        print("ฟังก์ชัน 'บันทึกข้อมูล' ถูกเรียกใช้งาน")
        # ในอนาคตสามารถเพิ่มโค้ดเพื่อดึงข้อมูลจาก Entry ต่างๆ มาบันทึกที่นี่

    def clear_data(self):
        """ฟังก์ชันตัวอย่างสำหรับล้างข้อมูลในช่องกรอก"""
        print("ฟังก์ชัน 'ล้างข้อมูล' ถูกเรียกใช้งาน")
        # ในอนาคตสามารถเพิ่มโค้ดเพื่อล้างค่าใน Entry ทั้งหมด


if __name__ == "__main__":
    root = tk.Tk()
    # --- 3. เรียกใช้ Theme ---
    # ใช้ 'light' หรือ 'dark'
    sv_ttk.set_theme("light")

    app = ModernDataEntryApp(root)
    root.mainloop()