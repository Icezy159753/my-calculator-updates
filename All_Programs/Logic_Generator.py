import tkinter as tk
# --- เปลี่ยนมาใช้ ttkbootstrap ---
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import re
import sys

class LogicGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Survey Logic Generator v8 (New UI)")
        
        # --- Center the window on the screen ---
        window_width = 900
        window_height = 750
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.root.minsize(800, 600)

        self.variable_map_case_insensitive = {}
        self.original_case_map = {}

        self._setup_styles()
        self._setup_ui()
        self._configure_highlighting_tags()

    def _setup_styles(self):
        """
        Set up custom styles. With ttkbootstrap, most styling is handled
        by the theme, so this is much simpler.
        """
        style = ttk.Style()
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
        
    def _setup_ui(self):
        # --- Layout using .grid() ---
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # --- Top configuration frame ---
        top_config_frame = ttk.Frame(self.root, padding=10)
        top_config_frame.grid(row=0, column=0, sticky="ew")
        top_config_frame.columnconfigure(0, weight=1)
        top_config_frame.columnconfigure(1, weight=1)

        # --- Section 1: Load Excel File ---
        file_frame = ttk.LabelFrame(top_config_frame, text=" 1. Load Excel File ", padding=10)
        file_frame.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="File Path:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state="readonly")
        file_entry.grid(row=0, column=1, sticky="ew")
        # --- ใช้ bootstyle ของ ttkbootstrap ---
        load_button = ttk.Button(file_frame, text="Browse & Load Items", command=self.load_excel_file, bootstyle="success")
        load_button.grid(row=0, column=2, sticky="e", padx=(5, 0))

        # --- Section 2: Column Settings ---
        config_frame = ttk.LabelFrame(top_config_frame, text=" 2. Excel Column Config ", padding=10)
        config_frame.grid(row=1, column=0, sticky="ew", pady=10, padx=(0, 5))
        
        self.marker_col_var = tk.StringVar(value="Segment")
        self.id_col_var = tk.StringVar(value="ID")
        self.type_col_var = tk.StringVar(value="Type")

        ttk.Label(config_frame, text="Segment Col:").grid(row=0, column=0, padx=(0,5), pady=5, sticky='w')
        ttk.Entry(config_frame, textvariable=self.marker_col_var, width=15).grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(config_frame, text="Variable ID Col:").grid(row=0, column=2, padx=(10,5), pady=5, sticky='w')
        ttk.Entry(config_frame, textvariable=self.id_col_var, width=15).grid(row=0, column=3, padx=5, pady=5, sticky='ew')
        ttk.Label(config_frame, text="Type Col:").grid(row=0, column=4, padx=(10,5), pady=5, sticky='w')
        ttk.Entry(config_frame, textvariable=self.type_col_var, width=15).grid(row=0, column=5, padx=5, pady=5, sticky='ew')
        config_frame.columnconfigure((1,3,5), weight=1)

        # --- Section 3: Select Identifiers ---
        id_selection_frame = ttk.LabelFrame(top_config_frame, text=" 3. Select Project Identifiers ", padding=10)
        id_selection_frame.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=(5, 0), pady=(0, 10))
        id_selection_frame.columnconfigure((1,3,5), weight=1)

        ttk.Label(id_selection_frame, text="ID 1 (Req):").grid(row=0, column=0, padx=(0,5), pady=5)
        self.id_var1 = tk.StringVar(value="KEY")
        self.id_var1_combo = ttk.Combobox(id_selection_frame, textvariable=self.id_var1, state="readonly", width=15)
        self.id_var1_combo.grid(row=0, column=1, sticky="ew", pady=5)

        ttk.Label(id_selection_frame, text="ID 2 (Opt):").grid(row=0, column=2, padx=(10,5), pady=5)
        self.id_var2 = tk.StringVar(value="UID")
        self.id_var2_combo = ttk.Combobox(id_selection_frame, textvariable=self.id_var2, state="readonly", width=15)
        self.id_var2_combo.grid(row=0, column=3, sticky="ew", pady=5)

        ttk.Label(id_selection_frame, text="ID 3 (Opt):").grid(row=0, column=4, padx=(10,5), pady=5)
        self.id_var3 = tk.StringVar(value="--- NONE ---")
        self.id_var3_combo = ttk.Combobox(id_selection_frame, textvariable=self.id_var3, state="readonly", width=15)
        self.id_var3_combo.grid(row=0, column=5, sticky="ew", pady=5)

        # --- PanedWindow for resizable central area ---
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # --- Left Pane: Item List ---
        item_list_frame = ttk.LabelFrame(main_pane, text=" 4. Item List (Double-click to add) ", padding=10)
        main_pane.add(item_list_frame, weight=1)
        item_list_frame.columnconfigure(0, weight=1)
        item_list_frame.rowconfigure(0, weight=1)

        list_scrollbar = ttk.Scrollbar(item_list_frame, orient=tk.VERTICAL)
        self.item_listbox = tk.Listbox(item_list_frame, yscrollcommand=list_scrollbar.set, font=("Courier New", 10), width=30)
        list_scrollbar.config(command=self.item_listbox.yview)
        self.item_listbox.grid(row=0, column=0, sticky="nsew")
        list_scrollbar.grid(row=0, column=1, sticky="ns")
        self.item_listbox.bind("<Double-1>", self.on_item_double_click)
        
        # --- Right Pane: Logic and Output ---
        logic_area_frame = ttk.Frame(main_pane, padding=10)
        main_pane.add(logic_area_frame, weight=3)
        logic_area_frame.columnconfigure(0, weight=1)
        logic_area_frame.rowconfigure(1, weight=2)
        logic_area_frame.rowconfigure(4, weight=3)

        # Logic Input Section
        ttk.Label(logic_area_frame, text="5. Logic Input (one condition per line):", style="Header.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 5))
        input_text_frame = ttk.Frame(logic_area_frame)
        input_text_frame.grid(row=1, column=0, sticky="nsew")
        input_text_frame.rowconfigure(0, weight=1)
        input_text_frame.columnconfigure(0, weight=1)
        input_scrollbar = ttk.Scrollbar(input_text_frame)
        self.logic_input_text = tk.Text(input_text_frame, height=10, font=("Courier New", 11), undo=True, yscrollcommand=input_scrollbar.set, wrap="word", relief="solid", borderwidth=1)
        input_scrollbar.config(command=self.logic_input_text.yview)
        self.logic_input_text.grid(row=0, column=0, sticky="nsew")
        input_scrollbar.grid(row=0, column=1, sticky="ns")
        self.logic_input_text.bind("<KeyRelease>", self._highlight_syntax)

        # --- ใช้ bootstyle ของ ttkbootstrap ---
        generate_button = ttk.Button(logic_area_frame, text="Generate & Append Syntax", command=self.generate_all_syntax, bootstyle="info")
        generate_button.grid(row=2, column=0, sticky="ew", pady=10)

        # Final Output Section
        output_header_frame = ttk.Frame(logic_area_frame)
        output_header_frame.grid(row=3, column=0, sticky="ew", pady=(10, 5))
        output_header_frame.columnconfigure(0, weight=1)
        ttk.Label(output_header_frame, text="6. Final Syntax:", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        # --- ใช้ bootstyle ของ ttkbootstrap ---
        ttk.Button(output_header_frame, text="Copy All", command=self.copy_to_clipboard, bootstyle="primary").grid(row=0, column=1, sticky="e", padx=5)
        ttk.Button(output_header_frame, text="Clear All", command=self.clear_final_output, bootstyle="danger").grid(row=0, column=2, sticky="e")
        ttk.Button(output_header_frame, text="Delete Selected Line", command=self.delete_selected_line, bootstyle="danger").grid(row=0, column=3, sticky="e", padx=(5,0))

        output_text_frame = ttk.Frame(logic_area_frame)
        output_text_frame.grid(row=4, column=0, sticky="nsew")
        output_text_frame.rowconfigure(0, weight=1)
        output_text_frame.columnconfigure(0, weight=1)
        output_scrollbar = ttk.Scrollbar(output_text_frame)
        self.final_syntax_text = tk.Text(output_text_frame, height=15, font=("Courier New", 11), yscrollcommand=output_scrollbar.set, wrap="word", relief="solid", borderwidth=1)
        output_scrollbar.config(command=self.final_syntax_text.yview)
        self.final_syntax_text.grid(row=0, column=0, sticky="nsew")
        output_scrollbar.grid(row=0, column=1, sticky="ns")


    def clear_final_output(self):
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all generated syntax?"):
            self.final_syntax_text.delete("1.0", tk.END)

    def _configure_highlighting_tags(self):
        text_widget = self.logic_input_text
        text_widget.tag_configure("variable", foreground="royal blue") 
        text_widget.tag_configure("operator", foreground="firebrick") 
        text_widget.tag_configure("keyword", foreground="purple")
        text_widget.tag_configure("number", foreground="forest green")
        text_widget.tag_configure("error", background="light pink")
        
        final_text_widget = self.final_syntax_text
        final_text_widget.tag_configure("statement", foreground="purple", font=("Courier New", 11, "bold"))
        final_text_widget.tag_configure("elist", foreground="dark cyan")
        final_text_widget.tag_configure("variable", foreground="royal blue")
        final_text_widget.tag_configure("operator", foreground="firebrick")
        final_text_widget.tag_configure("ma_func", foreground="dark goldenrod")

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
        variable_name = selected_text.split()[0] if selected_text.split() else ""
        if variable_name:
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
                missing_cols = [c for c in [marker_col, id_col, type_col] if c not in df.columns]
                messagebox.showerror("Error", f"Required column(s) not found in the Excel file: {', '.join(missing_cols)}")
                return

            self.variable_map_case_insensitive.clear()
            self.original_case_map.clear()
            self.item_listbox.delete(0, tk.END)
            
            for _, row in df.iterrows():
                var_id, segment = str(row[id_col]).strip(), str(row[marker_col]).strip()
                if segment and var_id and not var_id.isnumeric():
                    var_type = str(row[type_col]).strip().upper() or 'TEXT'
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

        next_error_num = 1
        try:
            existing_syntax = self.final_syntax_text.get("1.0", tk.END)
            if existing_syntax.strip():
                found_numbers = re.findall(r"-\s*(\d+)'", existing_syntax)
                if found_numbers:
                    last_max_num = max(int(n) for n in found_numbers)
                    next_error_num = last_max_num + 1
        except Exception as e:
            print(f"Could not determine next error number, defaulting to 1. Error: {e}")
            next_error_num = 1
        
        error_num_start = simpledialog.askinteger(
            "Starting Error Number", 
            "Enter the starting error number:", 
            initialvalue=next_error_num,
            parent=self.root
        )
        
        if error_num_start is None: return

        generated_syntaxes, current_error_num = [], error_num_start
        for i, line in enumerate(all_lines, 1):
            if not line.strip(): continue
            
            try:
                final_syntax, error_message = self.generate_single_syntax(line.strip(), current_error_num)
            except ValueError as ve:
                final_syntax, error_message = None, str(ve)

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

    def _transform_single_condition(self, condition_part, vars_used_in_logic):
            match = re.match(r'^\s*([a-zA-Z_][a-zA-Z0-9_]*(?:\(\d+\))?)\s*([=^<>]+)\s*(.*)\s*$', condition_part)
            if not match:
                raise ValueError(f"Invalid condition format: '{condition_part}'")

            var, op, val = match.group(1), match.group(2), match.group(3).strip()
            var_lower = var.lower()
            var_original_case = self.original_case_map.get(var_lower, var.upper())
            var_type = self.variable_map_case_insensitive.get(var_lower, {}).get("type", "UNKNOWN")

            vars_used_in_logic.append(var_original_case)

            if val.lower() in ['roff', 'ron', 'off1', 'off2']:
                return f"{var_original_case}{op}{val.upper()}"
            
            if "SA" in var_type:
                expanded_values = self._expand_sa_values(val)
                if expanded_values is None:
                    raise ValueError(f"Invalid value format '{val}' for SA variable '{var_original_case}'.")
                
                if op == '=':
                    if len(expanded_values) == 1:
                        return f"{var_original_case}={expanded_values[0]}"
                    else:
                        conditions = [f"{var_original_case}={v}" for v in expanded_values]
                        return f"({' OR '.join(conditions)})"
                elif op in ['^=', '<>']:
                    if len(expanded_values) == 1:
                        return f"{var_original_case}{op}{expanded_values[0]}"
                    else:
                        conditions = [f"{var_original_case}{op}{v}" for v in expanded_values]
                        return f"({' AND '.join(conditions)})"
                else:
                    return f"{var_original_case}{op}{val}"

            if "MA" in var_type:
                if re.search(r'\d', val):
                    return f"{var_original_case}{op}MFT({val})"
                else:
                    return f"{var_original_case}{op}{val}"
            
            return f"{var_original_case}{op}{val}"

    def _parse_logical_expression(self, expression, vars_used_in_logic):
            expression = expression.strip()
            if expression.startswith('(') and expression.endswith(')'):
                balance = 0
                is_wrapped = True
                for i, char in enumerate(expression[1:-1]):
                    if char == '(': balance += 1
                    elif char == ')': balance -= 1
                    if balance < 0:
                        is_wrapped = False
                        break
                if is_wrapped and balance == 0:
                    return self._parse_logical_expression(expression[1:-1], vars_used_in_logic)

            balance = 0
            parts = []
            last_split_idx = 0
            for i, char in enumerate(expression):
                if char == '(': balance += 1
                elif char == ')': balance -= 1
                elif char == '|' and balance == 0:
                    parts.append(expression[last_split_idx:i])
                    last_split_idx = i + 1
            
            if parts:
                parts.append(expression[last_split_idx:])
                parsed_parts = [self._parse_logical_expression(p, vars_used_in_logic) for p in parts]
                return f"({' OR '.join(parsed_parts)})"

            balance = 0
            parts = []
            last_split_idx = 0
            for i, char in enumerate(expression):
                if char == '(': balance += 1
                elif char == ')': balance -= 1
                elif char == '&' and balance == 0:
                    parts.append(expression[last_split_idx:i])
                    last_split_idx = i + 1

            if parts:
                parts.append(expression[last_split_idx:])
                parsed_parts = [self._parse_logical_expression(p, vars_used_in_logic) for p in parts]
                return f"({' AND '.join(parsed_parts)})"

            return self._transform_single_condition(expression, vars_used_in_logic)

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

                vars_used_in_logic = []
                final_condition_string = self._parse_logical_expression(user_input, vars_used_in_logic)
                
                if final_condition_string.startswith('(') and final_condition_string.endswith(')'):
                    final_condition_string = final_condition_string[1:-1]

                if not vars_used_in_logic:
                    return None, "No valid variables found in the logic condition."

                first_var_info = self.variable_map_case_insensitive.get(vars_used_in_logic[0].lower(), {})
                is_loop_base_var = first_var_info.get("is_loop_base", False) and '(' not in vars_used_in_logic[0]

                elist_id_vars = [id_var1]
                if id_var2 and id_var2 != "--- NONE ---": elist_id_vars.append(id_var2)
                if id_var3 and id_var3 != "--- NONE ---": elist_id_vars.append(id_var3)

                real_vars_for_elist = list(set(vars_used_in_logic))
                all_elist_vars = elist_id_vars + sorted(list(set(real_vars_for_elist) - set(elist_id_vars)))
                elist_vars = ','.join(all_elist_vars)
                
                reconstructed_input = user_input.replace("'", "\\'")
                elist_message = f"IF {reconstructed_input} - {error_num}"
                elist_part = f"ELIST ('{elist_message}', {elist_vars})"
                
                final_syntax = f"DO IF {final_condition_string} THEN {elist_part} FI OD" if is_loop_base_var else f"IF {final_condition_string} THEN {elist_part} FI"
                
                return final_syntax, None

            except ValueError as ve:
                return None, str(ve)
            except Exception as e:
                return None, f"An unexpected parsing error occurred: {e}"

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
        except tk.TclError: 
            pass

def run_this_app(working_dir=None):
    """
    Main function to create and run the LogicGeneratorApp.
    """
    try:
        # --- ใช้ ttkbootstrap.Window และเลือก theme, เช่น 'litera', 'superhero', 'flatly' ---
        root = ttk.Window(themename="litera")
        app = LogicGeneratorApp(root)
        root.mainloop()
    except Exception as e:
        print(f"ERROR: An error occurred during app execution: {e}")
        if 'root' not in locals() or not root.winfo_exists():
            root_temp = tk.Tk()
            root_temp.withdraw()
            messagebox.showerror("Application Error", f"An unexpected error occurred:\n{e}", parent=root_temp)
            root_temp.destroy()
        else:
            messagebox.showerror("Application Error", f"An unexpected error occurred:\n{e}", parent=root)
        sys.exit(f"Error running the app: {e}")



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