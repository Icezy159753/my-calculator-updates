import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import re

class LogicGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Survey Logic Generator v6")
        self.root.geometry("900x750")

        self.variable_map_case_insensitive = {}
        self.original_case_map = {}

        self._setup_ui()
        self._configure_highlighting_tags()

    def _setup_ui(self):
        # ... (This function remains unchanged)
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill="x", padx=5, pady=5)
        self.file_frame = tk.Frame(top_frame, pady=5, bd=1, relief="groove")
        self.file_frame.pack(fill="x")
        tk.Label(self.file_frame, text=" Load Excel File ", font=("Arial", 9, "bold")).place(x=10, y=-10)
        tk.Label(self.file_frame, text="Excel File Path:").pack(side="left", padx=(10,5), pady=10)
        self.file_path_var = tk.StringVar()
        tk.Entry(self.file_frame, textvariable=self.file_path_var, state="readonly").pack(side="left", fill="x", expand=True, pady=10)
        tk.Button(self.file_frame, text="Load Items", command=self.load_excel_file).pack(side="left", padx=5, pady=10)
        self.config_frame = tk.Frame(top_frame, pady=5, bd=1, relief="groove")
        self.config_frame.pack(fill="x", pady=(5,0))
        tk.Label(self.config_frame, text=" Excel Column Config ", font=("Arial", 9, "bold")).place(x=10, y=-10)
        self.marker_col_var = tk.StringVar(value="Segment")
        self.id_col_var = tk.StringVar(value="ID")
        self.type_col_var = tk.StringVar(value="Type")
        tk.Label(self.config_frame, text="Segment Col:").pack(side="left", padx=(10,5), pady=10)
        tk.Entry(self.config_frame, textvariable=self.marker_col_var, width=12).pack(side="left", pady=10)
        tk.Label(self.config_frame, text="Variable ID Col:").pack(side="left", padx=(15,5), pady=10)
        tk.Entry(self.config_frame, textvariable=self.id_col_var, width=12).pack(side="left", pady=10)
        tk.Label(self.config_frame, text="Type Col:").pack(side="left", padx=(15,5), pady=10)
        tk.Entry(self.config_frame, textvariable=self.type_col_var, width=12).pack(side="left", pady=10)
        self.id_selection_frame = tk.Frame(self.root, pady=5, bd=1, relief="groove")
        self.id_selection_frame.pack(fill="x", padx=5, pady=5)
        tk.Label(self.id_selection_frame, text=" Select Project Identifiers ", font=("Arial", 9, "bold")).place(x=10, y=-10)
        tk.Label(self.id_selection_frame, text="ID 1 (Req):").pack(side="left", padx=(10, 2), pady=10)
        self.id_var1 = tk.StringVar(value="KEY")
        self.id_var1_combo = ttk.Combobox(self.id_selection_frame, textvariable=self.id_var1, state="readonly", width=15)
        self.id_var1_combo.pack(side="left", pady=10)
        tk.Label(self.id_selection_frame, text="ID 2 (Opt):").pack(side="left", padx=(20, 2), pady=10)
        self.id_var2 = tk.StringVar(value="UID")
        self.id_var2_combo = ttk.Combobox(self.id_selection_frame, textvariable=self.id_var2, state="readonly", width=15)
        self.id_var2_combo.pack(side="left", pady=10)
        tk.Label(self.id_selection_frame, text="ID 3 (Opt):").pack(side="left", padx=(20, 2), pady=10)
        self.id_var3 = tk.StringVar(value="--- NONE ---")
        self.id_var3_combo = ttk.Combobox(self.id_selection_frame, textvariable=self.id_var3, state="readonly", width=15)
        self.id_var3_combo.pack(side="left", pady=10)
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.item_list_frame = tk.Frame(self.main_frame, bd=1, relief="sunken")
        self.item_list_frame.pack(side="left", fill="y", padx=(0, 5))
        tk.Label(self.item_list_frame, text="Item List", font=("Arial", 10, "bold")).pack()
        self.item_listbox = tk.Listbox(self.main_frame, width=30)
        self.item_listbox.pack(in_=self.item_list_frame, fill="y", expand=True)
        self.item_listbox.bind("<Double-1>", self.on_item_double_click)
        self.logic_area_frame = tk.Frame(self.main_frame)
        self.logic_area_frame.pack(side="left", fill="both", expand=True)
        tk.Label(self.logic_area_frame, text="Logic Input (one condition per line):", font=("Arial", 10, "bold")).pack(anchor="w")
        input_frame = tk.Frame(self.logic_area_frame)
        input_frame.pack(fill="x")
        input_scrollbar = tk.Scrollbar(input_frame)
        input_scrollbar.pack(side="right", fill="y")
        self.logic_input_text = tk.Text(input_frame, height=10, font=("Courier New", 10), undo=True, yscrollcommand=input_scrollbar.set)
        self.logic_input_text.pack(side="left", fill="x", expand=True)
        input_scrollbar.config(command=self.logic_input_text.yview)
        self.logic_input_text.bind("<KeyRelease>", self._highlight_syntax)
        tk.Button(self.logic_area_frame, text="Generate & Append", command=self.generate_all_syntax, pady=5).pack(pady=5)
        final_syntax_header_frame = tk.Frame(self.logic_area_frame)
        final_syntax_header_frame.pack(fill="x", pady=(10, 0))
        tk.Label(final_syntax_header_frame, text="Final Syntax:", font=("Arial", 10, "bold")).pack(side="left")
        tk.Button(final_syntax_header_frame, text="Clear All", command=self.clear_final_output).pack(side="right")
        tk.Button(final_syntax_header_frame, text="Copy All", command=self.copy_to_clipboard).pack(side="right", padx=5)
        output_frame = tk.Frame(self.logic_area_frame)
        output_frame.pack(fill="both", expand=True)
        output_scrollbar = tk.Scrollbar(output_frame)
        output_scrollbar.pack(side="right", fill="y")
        self.final_syntax_text = tk.Text(output_frame, height=15, font=("Courier New", 10), bd=1, relief="sunken", yscrollcommand=output_scrollbar.set)
        self.final_syntax_text.pack(side="left", fill="both", expand=True)
        output_scrollbar.config(command=self.final_syntax_text.yview)
        self.output_button_frame = tk.Frame(self.logic_area_frame)
        self.output_button_frame.pack(fill="x", pady=5)
        tk.Button(self.output_button_frame, text="Delete Selected Line", command=self.delete_selected_line).pack(side="right", padx=5)

    def clear_final_output(self):
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all generated syntax?"):
            self.final_syntax_text.delete("1.0", tk.END)

    def _configure_highlighting_tags(self):
        text_widget = self.logic_input_text
        text_widget.tag_configure("variable", foreground="#0000FF") 
        text_widget.tag_configure("operator", foreground="#FF0000") 
        text_widget.tag_configure("keyword", foreground="#A020F0")  
        text_widget.tag_configure("number", foreground="#008000")   
        text_widget.tag_configure("error", background="#FFC0CB")    
        final_text_widget = self.final_syntax_text
        final_text_widget.tag_configure("statement", foreground="#A020F0", font=("Courier New", 10, "bold")) 
        final_text_widget.tag_configure("elist", foreground="#008B8B")
        final_text_widget.tag_configure("variable", foreground="#0000FF") 
        final_text_widget.tag_configure("operator", foreground="#FF0000")
        final_text_widget.tag_configure("ma_func", foreground="#DAA520")

    def _highlight_syntax(self, event=None):
        widget = self.logic_input_text
        tags_to_clear = ["variable", "operator", "keyword", "number"]
        for tag in tags_to_clear:
            widget.tag_remove(tag, "1.0", tk.END)
        content = widget.get("1.0", tk.END)
        known_vars = sorted(self.original_case_map.keys(), key=len, reverse=True)
        if known_vars:
            var_pattern = r'\b(' + '|'.join(re.escape(v) for v in known_vars) + r')\b'
            self._apply_tag_to_pattern(widget, var_pattern, "variable", content, re.IGNORECASE)
        self._apply_tag_to_pattern(widget, r'\b(roff|ron|off1|off2)\b', "keyword", content, re.IGNORECASE)
        self._apply_tag_to_pattern(widget, r'([=^<>]+|&|\|)', "operator", content)
        self._apply_tag_to_pattern(widget, r'\b\d+([-,]\s*\d+)*\b', "number", content)

    def _apply_tag_to_pattern(self, widget, pattern, tag, content, flags=0):
        for match in re.finditer(pattern, content, flags):
            start = f"1.0 + {match.start()} chars"
            end = f"1.0 + {match.end()} chars"
            widget.tag_add(tag, start, end)

    def _highlight_final_syntax(self):
        widget = self.final_syntax_text
        content = widget.get("1.0", tk.END)
        for tag in ["statement", "elist", "variable", "operator", "ma_func"]:
             widget.tag_remove(tag, "1.0", tk.END)
        self._apply_tag_to_pattern(widget, r'\b(DO|IF|THEN|FI|OD)\b', "statement", content)
        self._apply_tag_to_pattern(widget, r'ELIST\s*\(.*?\)', "elist", content)
        self._apply_tag_to_pattern(widget, r'\b(MFT|ML)\(.*\)', "ma_func", content)
        self._apply_tag_to_pattern(widget, r'\b([a-zA-Z_][a-zA-Z0-9_]*(?:\(\d+\))?)\b', "variable", content)
        self._apply_tag_to_pattern(widget, r'[=^<>]+|AND|OR', "operator", content)

    def on_item_double_click(self, event):
        selection_indices = self.item_listbox.curselection()
        if not selection_indices: return
        selected_text = self.item_listbox.get(selection_indices[0])
        variable_name = selected_text.split()[0]
        self.logic_input_text.insert(tk.INSERT, variable_name + " ")
        self.logic_input_text.focus_set()
        self._highlight_syntax()

    def load_excel_file(self):
        file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path: return
        self.file_path_var.set(file_path)
        try:
            df = pd.read_excel(file_path, header=1, dtype=str).fillna('')
            marker_col, id_col, type_col = self.marker_col_var.get(), self.id_col_var.get(), self.type_col_var.get()
            if not all(c in df.columns for c in [marker_col, id_col, type_col]):
                 messagebox.showerror("Error", f"Required columns '{marker_col}', '{id_col}', or '{type_col}' not found.")
                 return
            self.variable_map_case_insensitive.clear()
            self.original_case_map.clear()
            self.item_listbox.delete(0, tk.END)
            for _, row in df.iterrows():
                var_id, segment = row[id_col].strip(), row[marker_col].strip()
                if segment and var_id and not var_id.isnumeric():
                    var_type = row[type_col].strip().upper() or 'TEXT'
                    self.variable_map_case_insensitive[var_id.lower()] = {"type": var_type, "is_loop_base": "LOOP" in var_type}
                    self.original_case_map[var_id.lower()] = var_id
                    self.item_listbox.insert(tk.END, f"{var_id:<15} {var_type}")
                elif not segment and var_id and '(' in var_id and ')' in var_id:
                    base_var_name_match = re.match(r'(\w+)\(', var_id)
                    sub_var_type = "UNKNOWN"
                    if base_var_name_match:
                        base_var_name = base_var_name_match.group(1).lower()
                        base_var_info = self.variable_map_case_insensitive.get(base_var_name, {})
                        base_var_type = base_var_info.get("type", "LOOP(UNKNOWN)")
                        sub_var_type = re.sub(r'LOOP\(|\)', '', base_var_type)
                    self.variable_map_case_insensitive[var_id.lower()] = {"type": sub_var_type, "is_loop_base": False}
                    self.original_case_map[var_id.lower()] = var_id
                    self.item_listbox.insert(tk.END, f"{var_id}")
            variable_names = list(self.original_case_map.values())
            none_option = ["--- NONE ---"]
            self.id_var1_combo['values'] = variable_names
            self.id_var2_combo['values'] = none_option + variable_names
            self.id_var3_combo['values'] = none_option + variable_names
            default_key = next((v for v in variable_names if "key" in v.lower() or "sbj" in v.lower()), variable_names[0] if variable_names else "")
            remaining_vars = [v for v in variable_names if v.lower() != default_key.lower()]
            default_uid = next((v for v in remaining_vars if "uid" in v.lower() or "id" in v.lower()), "--- NONE ---")
            self.id_var1.set(default_key)
            self.id_var2.set(default_uid)
            self.id_var3.set("--- NONE ---")
            messagebox.showinfo("Success", f"{len(self.original_case_map)} items loaded successfully.")
            self._highlight_syntax()
        except Exception as e:
            messagebox.showerror("Error Loading File", f"An error occurred: {e}")
            
    def generate_all_syntax(self):
        self.logic_input_text.tag_remove("error", "1.0", tk.END)
        all_lines = self.logic_input_text.get("1.0", tk.END).strip().split('\n')
        if not any(all_lines):
            messagebox.showwarning("Input Missing", "Please enter at least one logic condition.")
            return
        error_num_start = simpledialog.askinteger("Starting Error Number", "Enter the starting error number:", initialvalue=1)
        if error_num_start is None: return
        generated_syntaxes, current_error_num = [], error_num_start
        for i, line in enumerate(all_lines, 1):
            if not line.strip(): continue
            final_syntax, error_message = self.generate_single_syntax(line.strip(), current_error_num)
            if error_message:
                line_start, line_end = f"{i}.0", f"{i}.end"
                self.logic_input_text.tag_add("error", line_start, line_end)
                self.logic_input_text.see(line_start)
                messagebox.showerror("Syntax Error", f"Error on line {i}:\n\n{error_message}")
                return
            if final_syntax:
                generated_syntaxes.append(final_syntax)
                current_error_num += 1
        if generated_syntaxes:
            if self.final_syntax_text.get("1.0", tk.END).strip():
                self.final_syntax_text.insert(tk.END, "\n")
            self.final_syntax_text.insert(tk.END, "\n".join(generated_syntaxes))
            self.logic_input_text.delete("1.0", tk.END)
            self._highlight_final_syntax()
            self.final_syntax_text.see(tk.END)
            messagebox.showinfo("Success", f"Appended {len(generated_syntaxes)} new syntax lines successfully.")
    
    def _expand_sa_values(self, value_str):
        final_values = set()
        parts = value_str.split(',')
        for part in parts:
            part = part.strip()
            if not part: continue
            if '-' in part:
                range_match = re.fullmatch(r'(\d+)-(\d+)', part)
                if not range_match: return None
                start, end = map(int, range_match.groups())
                if start > end: return None
                final_values.update(map(str, range(start, end + 1)))
            elif part.isnumeric():
                final_values.add(part)
            else:
                return None
        return sorted(list(final_values), key=int)

    def generate_single_syntax(self, user_input, error_num):
        if not self.variable_map_case_insensitive:
            return None, "Please load an item list first."
        id_var1 = self.id_var1.get()
        if not id_var1 or id_var1 == "--- NONE ---":
            return None, "Please select the required ID Variable 1."
        try:
            all_potential_vars = set(re.findall(r'\b[a-zA-Z_][a-zA-Z0-9_]*(?:\(\d+\))?\b', user_input.lower()))
            id_var2 = self.id_var2.get()
            id_var3 = self.id_var3.get()
            special_words = ['roff', 'ron', 'off1', 'off2', id_var1.lower()]
            if id_var2 and id_var2 != "--- NONE ---": special_words.append(id_var2.lower())
            if id_var3 and id_var3 != "--- NONE ---": special_words.append(id_var3.lower())
            invalid_vars = [var for var in all_potential_vars if var not in self.variable_map_case_insensitive and var not in special_words and not var.isnumeric()]
            if invalid_vars: 
                return None, f"Variable(s) not found: {', '.join(invalid_vars)}"

            transformed_parts = []
            vars_used_in_logic = [] 

            for part in re.split(r'\s*([&|])\s*', user_input):
                part = part.strip()
                if not part: continue
                if part == '&': transformed_parts.append(" AND "); continue
                if part == '|': transformed_parts.append(" OR "); continue
                
                match = re.match(r'^\s*([a-zA-Z_][a-zA-Z0-9_]*(?:\(\d+\))?)\s*([=^<>]+)\s*(.*)\s*$', part)
                if not match: continue
                
                var, op, val = match.group(1), match.group(2), match.group(3).strip()
                
                var_lower = var.lower()
                var_original_case = self.original_case_map.get(var_lower, var.upper())
                var_type = self.variable_map_case_insensitive.get(var_lower, {}).get("type", "UNKNOWN")

                vars_used_in_logic.append(var_original_case)

                # --- START: THE ONLY CRITICAL CHANGE IN THIS VERSION ---
                if val.lower() in ['roff', 'ron', 'off1', 'off2']:
                    transformed_parts.append(f"{var_original_case}{op}{val.upper()}")
                
                elif "SA" in var_type:
                    expanded_values = self._expand_sa_values(val)
                    if expanded_values is None:
                        return None, f"Invalid value format '{val}' for SA variable '{var_original_case}'. Use numbers, commas, and hyphens only (e.g., '1-3,5')."
                    
                    if op == '=':
                        if len(expanded_values) == 1:
                            transformed_parts.append(f"{var_original_case}={expanded_values[0]}")
                        else:
                            or_conditions = [f"{var_original_case}={v}" for v in expanded_values]
                            transformed_parts.append(f"({' OR '.join(or_conditions)})")
                    
                    elif op in ['^=', '<>']:
                        if len(expanded_values) == 1:
                            transformed_parts.append(f"{var_original_case}{op}{expanded_values[0]}")
                        else:
                            and_conditions = [f"{var_original_case}{op}{v}" for v in expanded_values]
                            transformed_parts.append(f"({' AND '.join(and_conditions)})")
                    
                    else: # For operators like > or < on SA variables
                        transformed_parts.append(f"{var_original_case}{op}{val}")

                elif "MA" in var_type:
                    if re.search(r'\d', val):
                        transformed_parts.append(f"{var_original_case}{op}MFT({val})")
                    else:
                        transformed_parts.append(f"{var_original_case}{op}{val}")
                else: 
                    transformed_parts.append(f"{var_original_case}{op}{val}")
                # --- END: CRITICAL CHANGE ---

            final_condition_string = "".join(transformed_parts)
            
            first_var_info = self.variable_map_case_insensitive.get(vars_used_in_logic[0].lower(), {})
            is_loop_base_var = first_var_info.get("is_loop_base", False) and '(' not in vars_used_in_logic[0]

            elist_id_vars = [id_var1]
            if id_var2 and id_var2 != "--- NONE ---": elist_id_vars.append(id_var2)
            if id_var3 and id_var3 != "--- NONE ---": elist_id_vars.append(id_var3)

            real_vars_for_elist = list(set(vars_used_in_logic))
            all_elist_vars = elist_id_vars + sorted(list(set(real_vars_for_elist) - set(elist_id_vars)))
            elist_vars = ','.join(all_elist_vars)
            
            reconstructed_input = final_condition_string.replace(' AND ', ' & ').replace(' OR ', ' | ').replace("'", "\\'")
            elist_message = f"IF {reconstructed_input} - {error_num}"
            elist_part = f"ELIST ('{elist_message}', {elist_vars})"
            
            final_syntax = f"DO IF {final_condition_string} THEN {elist_part} FI OD" if is_loop_base_var else f"IF {final_condition_string} THEN {elist_part} FI"
            
            return final_syntax, None
        except Exception as e:
            return None, f"An unexpected error occurred: {e}"

    def copy_to_clipboard(self):
        content = self.final_syntax_text.get("1.0", tk.END)
        if not content.strip(): return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        messagebox.showinfo("Copied", "Syntax has been copied to the clipboard.")
        
    def delete_selected_line(self):
        try:
            line_start = self.final_syntax_text.index("insert linestart")
            line_end = self.final_syntax_text.index("insert lineend")
            self.final_syntax_text.delete(line_start, line_end + "+1c")
        except tk.TclError: pass



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
        root = tk.Tk()
        app = LogicGeneratorApp(root)
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