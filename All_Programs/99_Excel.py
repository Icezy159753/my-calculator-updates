import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, Canvas, Frame
import pandas as pd # ต้องมี pandas และ engine สำหรับอ่าน Excel (เช่น openpyxl) ติดตั้งไว้
import random
from collections import defaultdict
import math # ไม่ได้ใช้ แต่คงไว้ก่อน
import os
import ast # ใช้สำหรับ Parse Format ('Val1', 'Val2'): Count อย่างปลอดภัย
import sys
import itertools # <--- เพิ่มบรรทัดนี้


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        # getattr checks if sys has _MEIPASS attribute
        base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    except Exception:
         # Fallback to script directory if _MEIPASS check fails unexpectedly
        base_path = os.path.abspath(os.path.dirname(__file__))

    return os.path.join(base_path, relative_path)
# --- ฟังก์ชัน Flexible Quota Sampling (ไม่ต้องแก้ไข) ---
def flexible_quota_sampling(population_df, quota_definitions, id_col='id'):
    """
    ทำการสุ่มตัวอย่างแบบ Greedy โดยพิจารณา Quota 1 หรือหลายชุดพร้อมกัน

    Args:
        population_df (pd.DataFrame): DataFrame ประชากรทั้งหมด
        quota_definitions (list): List ของ tuples, โดยแต่ละ tuple คือ (dimensions, targets_dict)
                                  เช่น [ (['S8','Quota'], {('กรุงเทพฯ', 'LINE MAN Users x BKK'):300}),
                                        (['ses', 'region'], {('AB', 'GBKK'): 60}) , ... ]
        id_col (str): ชื่อคอลัมน์ ID

    Returns:
        tuple: (selected_ids, list_of_final_counts, list_of_unmet_quotas)
    """
    # --- การตรวจสอบ Input เบื้องต้น ---
    if id_col not in population_df.columns:
        raise ValueError(f"population_df must contain an '{id_col}' column.")

    num_quotas = len(quota_definitions)
    if num_quotas == 0:
        raise ValueError("No quota definitions provided.")

    all_dims = set()
    for i, (dims, targets) in enumerate(quota_definitions):
        if not isinstance(dims, list): raise ValueError(f"Quota Set {i+1} Dimensions must be a list, got: {dims}")
        if not isinstance(targets, dict): raise ValueError(f"Quota Set {i+1} Targets must be a dict, got: {type(targets)}")
        all_dims.update(dims)

    for dim in all_dims:
        if dim not in population_df.columns:
                 raise ValueError(f"Dimension column '{dim}' not found in population_df. Available columns: {list(population_df.columns)}")
        # Consider checking for mixed types or converting columns to string if values can be numeric/string
        # population_df[dim] = population_df[dim].astype(str) # Example conversion
        if population_df[dim].isnull().any():
            print(f"Info: Dimension column '{dim}' contains NaN/None values. Rows with NaN in this dimension will be excluded during key generation.")

    # --- การเตรียมข้อมูลภายในฟังก์ชัน ---
    available_candidates_df = population_df.copy()
    # Ensure ID column is string for reliable set operations
    try:
        available_candidates_df[id_col] = available_candidates_df[id_col].astype(str)
    except Exception as e:
        raise ValueError(f"Could not convert ID column '{id_col}' to string: {e}")

    available_ids = set(available_candidates_df[id_col])
    selected_ids = set()
    current_counts = [defaultdict(int) for _ in range(num_quotas)]
    quota_keys_cols = [f'quota_key_{i}' for i in range(num_quotas)]

    # --- ฟังก์ชัน Helper สร้าง Key ---
    def get_quota_key(row, dimensions):
        try:
            # Ensure all dimension values are treated as strings for consistent tuple keys
            # This helps if a column mixes numbers and strings that should be treated the same
            values = [str(row[dim]) for dim in dimensions]
            # Check for actual missing values *after* potential conversion (though isnull check earlier is better)
            if any(pd.isna(row[dim]) for dim in dimensions): return None
            return tuple(values)
        except KeyError as e: print(f"Warn: Key gen error - Missing dimension {e} for row id {row.get(id_col, 'N/A')}"); return None
        except Exception as e: print(f"Warn: Key gen error: {e} for row id {row.get(id_col, 'N/A')}"); return None

    # --- สร้าง Quota Keys และกรอง ---
    print("Generating quota keys...")
    valid_rows = True # Start assuming rows are valid
    for i, (dims, _) in enumerate(quota_definitions):
        key_col_name = f'quota_key_{i}'
        available_candidates_df[key_col_name] = available_candidates_df.apply(
            lambda row: get_quota_key(row, dims), axis=1
        )
        # Check if any keys failed to generate for this quota set
        if available_candidates_df[key_col_name].isnull().all():
            print(f"ERROR: All quota keys for Set {i+1} (dims: {dims}) resulted in None. Check column data and types.")
            valid_rows = False # Mark as invalid if a whole set fails

    initial_count = len(available_candidates_df)
    # Drop rows where *any* of the quota keys are None
    available_candidates_df = available_candidates_df.dropna(subset=quota_keys_cols)
    filtered_count = len(available_candidates_df)

    if initial_count > filtered_count:
        print(f"Info: Excluded {initial_count - filtered_count} candidates due to missing data or key generation errors.")
        available_ids = set(available_candidates_df[id_col]) # Update available IDs

    if filtered_count == 0 or not valid_rows:
        print("Warning: No valid candidates remaining after checking quota dimensions/keys.")
        unmet_quotas_list = [targets.copy() for _, targets in quota_definitions]
        return set(), [defaultdict(int) for _ in range(num_quotas)], unmet_quotas_list

    # --- ตรวจสอบความพร้อมเบื้องต้น (Optional) ---
    print("\n--- Initial Candidate Availability Check ---")
    possible_to_meet_all = True
    for i, (dims, targets) in enumerate(quota_definitions):
        print(f"Quota Set {i+1} Availability vs Target (Dims: {dims}):")
        try:
            # Ensure keys in targets are tuples of strings for matching
            stringified_targets = {tuple(map(str, k)): v for k, v in targets.items()}
            actual_counts = available_candidates_df[f'quota_key_{i}'].value_counts().to_dict() # Keys here are already tuples of strings from get_quota_key
            for key, target in stringified_targets.items():
                available = actual_counts.get(key, 0)
                print(f"- {key}: Target={target}, Available={available}", end="")
                if available < target: print(" <-- WARNING: Insufficient!"); possible_to_meet_all = False
                else: print()
        except Exception as e:
            print(f"Error during availability check for Quota Set {i+1}: {e}")
            possible_to_meet_all = False # Assume not possible if check fails

    if not possible_to_meet_all: print("\nWarning: Based on initial counts, not all targets can be met.")
    else: print("\nInfo: Initial counts suggest enough candidates exist for each cell (competition still applies).")
    print("--------------------------------------------\n")

    # --- ขั้นตอนหลักของ Greedy Algorithm ---
    print("Starting Greedy Selection Process...")
    # Use target size from the first quota set as the overall goal
    try:
        if not quota_definitions[0][1]: # Check if first target dict is empty
             raise ValueError("Quota Set 1 has no targets defined.")
        total_target_size = sum(quota_definitions[0][1].values())
        if total_target_size <= 0: raise ValueError("Total target size from Quota Set 1 must be positive.")
    except Exception as e:
        raise ValueError(f"Could not calculate total target size from Quota Set 1: {e}")

    max_iterations = total_target_size * 3 # Safety break (increased slightly)
    iteration_count = 0

    while len(selected_ids) < total_target_size and iteration_count < max_iterations :
        iteration_count += 1
        # 1. คำนวณ Need สำหรับทุกชุด
        needs_list = []
        total_needs_list = []
        try:
            for i, (_, targets) in enumerate(quota_definitions):
                # Ensure keys are tuples of strings for lookup against current_counts
                stringified_targets_loop = {tuple(map(str, k)): v for k, v in targets.items()}
                needs = {k: t - current_counts[i].get(k, 0) for k, t in stringified_targets_loop.items() if t - current_counts[i].get(k, 0) > 0}
                needs_list.append(needs)
                total_needs_list.append(sum(needs.values()))
        except Exception as e:
            print(f"ERROR calculating needs in iteration {iteration_count}: {e}")
            break # Stop if needs calculation fails

        # 2. ตรวจสอบเงื่อนไขหยุด
        if all(total_need == 0 for total_need in total_needs_list):
                 print(f"\nAll quota needs met at iteration {iteration_count}. Selected: {len(selected_ids)}.")
                 break
        if not available_ids:
            print(f"\nRan out of available candidates at iteration {iteration_count}. Selected: {len(selected_ids)}.")
            break

        # 3. ค้นหาผู้สมัครที่ดีที่สุด
        # Filter available candidates only once
        current_available_df = available_candidates_df[available_candidates_df[id_col].isin(available_ids)].copy()
        if current_available_df.empty:
            print(f"\nNo candidates remaining in the available pool at iteration {iteration_count}.")
            break # Should be caught by available_ids check, but safety

        # Calculate scores
        try:
            current_available_df['greedy_score'] = 0 # Initialize score column
            for i in range(num_quotas):
                # Map needs using the quota_key column (which are tuples of strings)
                need_map = needs_list[i]
                # Only add score if the key exists in the needs map (i.e., need > 0)
                current_available_df[f'need_{i}_score'] = current_available_df[f'quota_key_{i}'].map(need_map).fillna(0)
                current_available_df['greedy_score'] += current_available_df[f'need_{i}_score']
        except Exception as e:
                 print(f"ERROR calculating scores in iteration {iteration_count}: {e}")
                 break # Stop if score calculation fails

        potential_candidates = current_available_df[current_available_df['greedy_score'] > 0]
        if potential_candidates.empty:
            print(f"\nNo remaining candidates contribute to unmet quotas at iteration {iteration_count}.")
            break # Stop if no one can help

        max_need_score = potential_candidates['greedy_score'].max()
        top_candidates = potential_candidates[potential_candidates['greedy_score'] == max_need_score]

        # Select and update
        try:
            # Ensure there are candidates before sampling
            if top_candidates.empty:
                 print(f"Warn: Top candidates became empty unexpectedly at iteration {iteration_count}.")
                 break

            selected_row = top_candidates.sample(n=1).iloc[0]
            best_candidate_id = selected_row[id_col] # Already string
            selected_keys = [selected_row[f'quota_key_{i}'] for i in range(num_quotas)] # Keys are tuples of strings

            selected_ids.add(best_candidate_id)
            available_ids.remove(best_candidate_id)
            for i in range(num_quotas):
                 if selected_keys[i] is not None: # Make sure key is valid before incrementing
                    current_counts[i][selected_keys[i]] += 1 # Use the string-tuple key

        except ValueError as ve: # Handles empty top_candidates if sample fails, or other value errors
                 print(f"Warn: Sampling selection error ({ve}) at iteration {iteration_count}.")
                 break # Stop if selection fails
        except KeyError as ke: # Handle potential KeyError if a quota_key column is missing unexpectedly
                 print(f"Warn: Selection error - KeyError accessing key columns ({ke}) at iteration {iteration_count}.")
                 break
        except Exception as e:
                 print(f"Warn: General selection error: {e} at iteration {iteration_count}.")
                 break # Stop on other errors

    # --- Post Loop ---
    if iteration_count >= max_iterations:
        print(f"Warning: Reached maximum iterations ({max_iterations}). Selected: {len(selected_ids)}.")

    print(f"\nSelection process finished. Total selected: {len(selected_ids)}")
    unmet_quotas_list = []
    final_counts_list_str_keys = [] # Use stringified keys for return consistency
    try:
        for i, (_, targets) in enumerate(quota_definitions):
            stringified_targets_final = {tuple(map(str, k)): v for k, v in targets.items()}
            # Use .get(k, 0) for current counts in case a target key was never selected
            unmet = {k: t - current_counts[i].get(k, 0) for k, t in stringified_targets_final.items() if t - current_counts[i].get(k, 0) > 0}
            unmet_quotas_list.append(unmet)
            # Ensure final counts dict uses stringified tuple keys and includes all target keys
            final_counts_str = defaultdict(int)
            final_counts_str.update(current_counts[i]) # Update with actual counts
            for key_tuple_str in stringified_targets_final: # Ensure all target keys exist
                 final_counts_str.setdefault(key_tuple_str, 0)
            final_counts_list_str_keys.append(dict(final_counts_str)) # Convert defaultdict to dict for return

    except Exception as e:
        print(f"ERROR during final calculation of unmet quotas: {e}")
        # Return potentially incomplete results or raise error? For GUI, maybe return best effort.
        return selected_ids, [dict(c) for c in current_counts], unmet_quotas_list # Return non-stringified keys counts if error

    return selected_ids, final_counts_list_str_keys, unmet_quotas_list
# --- จบฟังก์ชัน flexible_quota_sampling ---


# --- คลาสสำหรับ GUI Application ---
class QuotaSamplerApp:
    # --- __init__ (ปรับข้อความปุ่ม Load) ---
    def __init__(self, master):
        self.master = master
        master.title("Program ตัด Quota V1") # Update title
        # +++++ เพิ่มบรรทัดนี้เพื่อตั้งค่าไอคอน +++++
        try:
            icon_path = resource_path("Cut.ico")
            master.iconbitmap(icon_path)
            print(f"DEBUG: Attempting to load icon from: {icon_path}") # Optional debug print
        except tk.TclError as e:
            print(f"Warning: Could not load icon 'Cut.ico'. Error: {e}")
        except Exception as e:
             print(f"Warning: An unexpected error occurred setting the icon: {e}")
        # ++++++++++++++++++++++++++++++++++++++++
        master.geometry("1300x800")

        self.file_path = tk.StringVar()
        self.loaded_df = None
        self.cleaned_df = None
        self.sampling_results = None
        self.id_col_target = 'id'
        self.num_quota_sets = 7

        self.quota_targets_data = [{} for _ in range(self.num_quota_sets)]

        # --- Widget storage ---
        self.quota_frames = []
        self.quota_enable_vars = []
        self.quota_enable_cbs = []
        self.quota_dim_listboxes = [None] * self.num_quota_sets # List to store Listbox widgets
        self.quota_dim_scrollbars = [None] * self.num_quota_sets # Store scrollbars too
        self.quota_dim_value_frames = [None] * self.num_quota_sets
        self.quota_dim_comboboxes = [[] for _ in range(self.num_quota_sets)]
        self.quota_target_count_entries = [None] * self.num_quota_sets
        self.quota_target_listboxes = [None] * self.num_quota_sets


        style = ttk.Style(); style.configure('TButton', padding=6, relief="flat", font=('Segoe UI', 10)); style.configure('TLabel', padding=2, font=('Segoe UI', 10)); style.configure('TEntry', padding=4, font=('Segoe UI', 10)); style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold')); style.configure('TCheckbutton', font=('Segoe UI', 10)); style.configure('TNotebook.Tab', padding=[12, 5], font=('Segoe UI', 10)); style.configure('TCombobox', padding=4, font=('Segoe UI', 10))
        main_frame = ttk.Frame(master, padding="10"); main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 1. Load Data ---
        load_frame = ttk.LabelFrame(main_frame, text="1. Load Data", padding="10"); load_frame.pack(fill=tk.X, pady=5, ipady=5)
        # ********* เปลี่ยนข้อความปุ่ม *********
        btn_load = ttk.Button(load_frame, text="Open Excel (.xlsx, .xls)", command=self.load_dataset); btn_load.pack(side=tk.LEFT, padx=5)
        self.lbl_file_path = ttk.Label(load_frame, textvariable=self.file_path, relief=tk.SUNKEN, width=80); self.lbl_file_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5); self.file_path.set("No file loaded.")

        # --- 2. Define Quotas (Using Notebook) ---
        quota_area_frame = ttk.LabelFrame(main_frame, text="2. กลุ่ม Quotaทั้งหมดสามารถเลือกได้", padding="10")
        quota_area_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.quota_notebook = ttk.Notebook(quota_area_frame)
        self.quota_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for i in range(self.num_quota_sets):
            q_frame = ttk.Frame(self.quota_notebook, padding="10")
            self.quota_notebook.add(q_frame, text=f"Quota Set {i+1}")
            self.quota_frames.append(q_frame)

            # --- Top Frame for Enable CB and Dimension Selection ---
            top_controls_frame = ttk.Frame(q_frame)
            top_controls_frame.pack(fill=tk.X, pady=(0, 5))

            # --- Enable/Disable Checkbox ---
            enable_var = tk.BooleanVar(value=(i == 0))
            self.quota_enable_vars.append(enable_var)
            if i > 0:
                cb_use_q = ttk.Checkbutton(top_controls_frame, text="Enable", variable=enable_var, command=lambda idx=i: self.toggle_quota_state(idx))
                cb_use_q.pack(side=tk.LEFT, anchor=tk.NW, padx=(0, 10)) # Anchor NW
                self.quota_enable_cbs.append(cb_use_q)
            else:
                self.quota_enable_cbs.append(None)

            # --- Dimensions Selection Area ---
            dim_select_frame = ttk.Frame(top_controls_frame)
            dim_select_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

            lbl_dims = ttk.Label(dim_select_frame, text="กรุณาเลือกข้อทั้งหมดในการตัด Quota:")
            lbl_dims.pack(anchor=tk.W)

            # === Use Listbox for Dimension Selection ===
            listbox_dim_frame = ttk.Frame(dim_select_frame) # Frame to hold listbox and scrollbar
            listbox_dim_frame.pack(fill=tk.X, expand=True)

            listbox_dim = tk.Listbox(listbox_dim_frame, height=4, selectmode=tk.EXTENDED, exportselection=False) # Height=4 rows
            listbox_dim.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.quota_dim_listboxes[i] = listbox_dim

            scrollbar_dim = ttk.Scrollbar(listbox_dim_frame, orient=tk.VERTICAL, command=listbox_dim.yview)
            scrollbar_dim.pack(side=tk.LEFT, fill=tk.Y)
            listbox_dim.config(yscrollcommand=scrollbar_dim.set)
            self.quota_dim_scrollbars[i] = scrollbar_dim
            # ===========================================

            # --- Load Dimensions Button (Now loads values for selected dims) ---
            btn_load_dims = ttk.Button(top_controls_frame, text="ยืนยัน\nข้อที่จะใช้ตัด", command=lambda idx=i: self.load_dimension_values(idx))
            btn_load_dims.pack(side=tk.LEFT, padx=5, anchor=tk.N) # Anchor N

            # --- Frame for Dynamic Dimension Value Comboboxes ---
            dim_values_outer_frame = ttk.LabelFrame(q_frame, text="เลือกข้อที่เป็นเงื่อนไขการตัด Quota", padding=5)
            dim_values_outer_frame.pack(fill=tk.X, pady=5)
            dim_values_frame = ttk.Frame(dim_values_outer_frame)
            dim_values_frame.pack(fill=tk.X)
            self.quota_dim_value_frames[i] = dim_values_frame

            # --- Frame for Adding Target Count ---
            add_target_frame = ttk.Frame(q_frame)
            add_target_frame.pack(fill=tk.X, pady=5)
            lbl_target_count = ttk.Label(add_target_frame, text="จำนวนN=")
            lbl_target_count.pack(side=tk.LEFT, padx=(10, 5))
            vcmd = (self.master.register(self.validate_positive_int), '%P')
            entry_target_count = ttk.Entry(add_target_frame, width=10, validate='key', validatecommand=vcmd)
            entry_target_count.pack(side=tk.LEFT, padx=5)
            self.quota_target_count_entries[i] = entry_target_count
            btn_add_target = ttk.Button(add_target_frame, text="เพิ่ม Quota ที่จะตัด", command=lambda idx=i: self.add_target_cell(idx))
            btn_add_target.pack(side=tk.LEFT, padx=10)

            # --- Frame for Displaying and Removing Targets ---
            display_remove_frame = ttk.LabelFrame(q_frame, text="แสดงกลุ่ม Quota ที่จะตัด", padding=5)
            display_remove_frame.pack(fill=tk.BOTH, expand=True, pady=5)
            listbox_target = tk.Listbox(display_remove_frame, height=8, font=('Courier New', 9), selectmode=tk.EXTENDED)
            listbox_target.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
            self.quota_target_listboxes[i] = listbox_target
            scrollbar_targets = ttk.Scrollbar(display_remove_frame, orient=tk.VERTICAL, command=listbox_target.yview)
            scrollbar_targets.pack(side=tk.LEFT, fill=tk.Y)
            listbox_target.config(yscrollcommand=scrollbar_targets.set)
            btn_remove_target = ttk.Button(display_remove_frame, text="ลบรายการที่เลือก", command=lambda idx=i: self.remove_target_cell(idx))
            btn_remove_target.pack(side=tk.LEFT, padx=5, pady=5, anchor=tk.N)

            # Initial state based on enable_var
            self.toggle_quota_state(i)

        # --- 3. Controls ---
        control_frame = ttk.Frame(main_frame, padding="5"); control_frame.pack(fill=tk.X, pady=5)
        self.btn_run = ttk.Button(control_frame, text="เริ่มตัด Quota!!", command=self.run_sampling, state=tk.DISABLED); self.btn_run.pack(side=tk.LEFT, padx=10)
        self.btn_export = ttk.Button(control_frame, text="Export Excel (.xlsx)", command=self.export_results, state=tk.DISABLED); self.btn_export.pack(side=tk.RIGHT, padx=10)

        # --- 4. Output ---
        output_frame = ttk.LabelFrame(main_frame, text="Output Log & Summary", padding="10"); output_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.txt_output = scrolledtext.ScrolledText(output_frame, height=12, wrap=tk.WORD, font=('Segoe UI', 9)); self.txt_output.pack(fill=tk.BOTH, expand=True)

        self.log_message("กรุณาโหลดRawdata Excel(.xlsx/.xls), เลือกข้อ+เงื่อนไขและจำนวน Sample Size")

        # --- จัดหน้าต่างให้อยู่กึ่งกลางจอ (ไม่ต้องแก้ไข) ---
        master.update_idletasks()
        window_width = master.winfo_width()
        window_height = master.winfo_height()
        if window_width <= 1 or window_height <= 1:
             try:
                  geom_parts = master.geometry().split('+')[0].split('x')
                  window_width = int(geom_parts[0])
                  window_height = int(geom_parts[1])
             except: window_width, window_height = 1300, 800
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        x = int((screen_width - window_width) / 2)
        y = int((screen_height - window_height) / 2)
        master.geometry(f"{window_width}x{window_height}+{x}+{y}")
        # ---------------------------------------------------


    # --- Validation for Target Count Entry (ไม่ต้องแก้ไข) ---
    def validate_positive_int(self, P):
        """Allow only empty string or positive integers."""
        if P == "": return True
        try:
            val = int(P); return val >= 0
        except ValueError: return False

    # --- Methods ---
    def log_message(self, message): self.txt_output.insert(tk.END, message + "\n"); self.txt_output.see(tk.END); self.master.update_idletasks()
    def clear_log(self): self.txt_output.delete('1.0', tk.END)

    # --- toggle_quota_state (ไม่ต้องแก้ไข) ---
    def toggle_quota_state(self, index):
        """Enable/disable widgets within a specific quota set tab based on its checkbox."""
        if index < 0 or index >= self.num_quota_sets: return

        is_enabled = self.quota_enable_vars[index].get()
        widget_state = tk.NORMAL if is_enabled or index == 0 else tk.DISABLED

        # --- Helper function to configure state ---
        def configure_widget_state(widget, state):
            if not widget or not widget.winfo_exists(): return
            try:
                if isinstance(widget, (tk.Listbox, ttk.Combobox, ttk.Entry, scrolledtext.ScrolledText, ttk.Button)):
                    widget.configure(state=state)
                elif isinstance(widget, ttk.Scrollbar):
                     widget.configure(state=tk.NORMAL) # Keep scrollbars active
            except (tk.TclError, AttributeError): pass
            except Exception as e: print(f"Error configuring state for {widget}: {e}")

        widgets_to_control = []
        try:
             # Dimension Listbox Area
             widgets_to_control.append(self.quota_dim_listboxes[index])
             widgets_to_control.append(self.quota_dim_scrollbars[index])
             top_frame = self.quota_frames[index].winfo_children()[0]
             load_button = top_frame.winfo_children()[-1]
             if isinstance(load_button, ttk.Button): widgets_to_control.append(load_button)

             # Dimension Value Comboboxes Area
             if self.quota_dim_value_frames[index]:
                  for widget in self.quota_dim_value_frames[index].winfo_children():
                      widgets_to_control.append(widget)

             # Add Target Area
             widgets_to_control.append(self.quota_target_count_entries[index])
             add_target_frame_widgets = self.quota_frames[index].winfo_children()[2].winfo_children()
             add_button = add_target_frame_widgets[-1]
             if isinstance(add_button, ttk.Button): widgets_to_control.append(add_button)
             if len(add_target_frame_widgets) > 1 and isinstance(add_target_frame_widgets[0], ttk.Label):
                 widgets_to_control.append(add_target_frame_widgets[0])

             # Defined Targets Area
             widgets_to_control.append(self.quota_target_listboxes[index])
             defined_targets_frame = self.quota_frames[index].winfo_children()[3]
             for defined_child in defined_targets_frame.winfo_children():
                  if isinstance(defined_child, (ttk.Scrollbar, ttk.Button)):
                       widgets_to_control.append(defined_child)
        except (IndexError, AttributeError) as e:
            print(f"Warn: Error finding widgets to toggle state for index {index}: {e}")


        final_state = widget_state
        if index == 0: final_state = tk.NORMAL

        for widget in widgets_to_control:
             if widget == self.quota_enable_cbs[index] and index > 0:
                 configure_widget_state(widget, tk.NORMAL)
             else:
                 configure_widget_state(widget, final_state)

        if index == 0:
             for widget in widgets_to_control:
                  configure_widget_state(widget, tk.NORMAL)


    # --- load_dataset (ปรับปรุงหลัก) ---
    # --- load_dataset (ปรับปรุงให้ใช้คอลัมน์แรกเป็น ID อัตโนมัติ) ---
    # --- load_dataset (ปรับปรุงให้ใช้คอลัมน์แรกเป็น ID อัตโนมัติ) ---
    def load_dataset(self):
        # ********* เปลี่ยน filetypes ให้รองรับ .xlsx และ .xls *********
        filepath = filedialog.askopenfilename(
            title="Select Excel Dataset File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if not filepath: return
        self.log_message(f"Loading data from Excel: {filepath}")

        # --- Reset internal state (ไม่ต้องแก้ไข) ---
        self.loaded_df = None
        self.cleaned_df = None
        self.sampling_results = None
        self.quota_targets_data = [{} for _ in range(self.num_quota_sets)]
        for i in range(self.num_quota_sets):
            self.clear_dimension_values(i)
            if self.quota_dim_listboxes[i]:
                try:
                    self.quota_dim_listboxes[i].config(state=tk.NORMAL)
                    self.quota_dim_listboxes[i].delete(0, tk.END)
                    self.quota_dim_listboxes[i].config(state=tk.DISABLED)
                except tk.TclError: pass
            if self.quota_target_listboxes[i]:
                self.update_target_display(i)
            if self.quota_target_count_entries[i]:
                try: self.quota_target_count_entries[i].delete(0, tk.END)
                except tk.TclError: pass
        self.btn_run['state'] = tk.DISABLED
        self.btn_export['state'] = tk.DISABLED
        # ----------------------------

        try:
            # --- Column Name Definitions (ชื่อคอลัมน์อื่นๆ ที่อาจต้อง Clean/Rename) ---
            # ID_COLUMN_ORIGINAL = 'SbjNum' # <<<--- ไม่ใช้แล้ว จะหาคอลัมน์แรกอัตโนมัติ
            REGION_COLUMN_ORIGINAL = 'S1_GROUP' # ยังคงใช้สำหรับ Rename/Clean อื่นๆ (ถ้ามี)
            AGE_COLUMN_ORIGINAL = 'S4_RANGE'  # ยังคงใช้สำหรับ Rename/Clean อื่นๆ (ถ้ามี)
            SES_COLUMN_ORIGINAL = 'S9_SESX'   # ยังคงใช้สำหรับ Rename/Clean อื่นๆ (ถ้ามี)
            self.id_col_target = 'id'         # ชื่อเป้าหมายของคอลัมน์ ID คือ 'id'
            ses_col_target = 'ses'            # ชื่อเป้าหมายของคอลัมน์ SES (ถ้ามี)

            # --- Read Data from Excel ---
            # อ่านข้อมูลโดยยังไม่ระบุ dtype ของ ID ล่วงหน้า
            try:
                self.loaded_df = pd.read_excel(filepath, sheet_name=0, engine=None)
                self.log_message(f"Read data from the first sheet of the Excel file.")
            except Exception as read_err:
                messagebox.showerror("Error Reading Excel", f"Could not read the Excel file.\nMake sure the file is valid and you have 'openpyxl' installed (`pip install openpyxl`).\n\nError details: {read_err}")
                self.log_message(f"Error reading Excel: {read_err}")
                self.file_path.set("Error loading file.")
                return

            temp_df = self.loaded_df.copy()
            self.log_message(f"Read {len(temp_df)} records.")

            # --- *** ระบุคอลัมน์แรก และเตรียมเป็น ID *** ---
            if temp_df.empty or len(temp_df.columns) == 0:
                raise ValueError("Loaded Excel sheet is empty or has no columns.")

            # ดึงชื่อคอลัมน์แรกสุด
            first_col_name = temp_df.columns[0]
            self.log_message(f"Using the first column ('{first_col_name}') as the ID column.")

            # ตรวจสอบว่าคอลัมน์แรกมีข้อมูลหรือไม่ (ป้องกัน Error ตอนแปลง type)
            if temp_df[first_col_name].isnull().all():
                self.log_message(f"Warning: The first column '{first_col_name}' contains only missing values.")
                # อาจจะตัดสินใจหยุดทำงาน หรือแค่เตือน แล้วปล่อยให้ขั้นตอน rename ทำงานต่อไป
                # raise ValueError(f"The first column '{first_col_name}' used as ID cannot be entirely empty.")

            # แปลงข้อมูลคอลัมน์แรกให้เป็น string (สำคัญมากสำหรับ ID)
            try:
                temp_df[first_col_name] = temp_df[first_col_name].astype(str)
                self.log_message(f"Converted data in column '{first_col_name}' to string type for ID.")
            except Exception as e_conv:
                # หากแปลงไม่ได้ อาจมีปัญหาข้อมูลที่ไม่คาดคิด
                raise ValueError(f"Could not convert data in the first column ('{first_col_name}') to string: {e_conv}")
            # --- *** สิ้นสุดการเตรียม ID *** ---

            # --- Rename Columns ---
            # สร้าง map สำหรับ rename โดยใช้ชื่อคอลัมน์แรกที่เจอ
            rename_map = { first_col_name: self.id_col_target }

            # เพิ่มการ rename อื่นๆ ที่ยังต้องการ (เช่น SES) เข้าไปใน map
            # ตรวจสอบก่อนว่าชื่อคอลัมน์ SES เดิม ไม่ใช่ชื่อเดียวกับคอลัมน์แรก (กรณีแปลกๆ)
            if SES_COLUMN_ORIGINAL in temp_df.columns and SES_COLUMN_ORIGINAL != first_col_name:
                rename_map[SES_COLUMN_ORIGINAL] = ses_col_target
            elif SES_COLUMN_ORIGINAL == first_col_name:
                self.log_message(f"Warning: The first column ('{first_col_name}') is also defined as the original SES column ('{SES_COLUMN_ORIGINAL}'). It will be renamed to '{self.id_col_target}'. The target SES name '{ses_col_target}' might not be created unless another column matches.")
            # เพิ่มการ rename อื่นๆ ตามต้องการ...
            # if REGION_COLUMN_ORIGINAL in temp_df.columns and REGION_COLUMN_ORIGINAL != first_col_name:
            #     rename_map[REGION_COLUMN_ORIGINAL] = 'region'

            # ทำการ rename เฉพาะคอลัมน์ที่มีอยู่จริงใน temp_df
            actual_renames = {k: v for k, v in rename_map.items() if k in temp_df.columns}
            if actual_renames:
                temp_df.rename(columns=actual_renames, inplace=True)
                self.log_message(f"Renamed columns: {actual_renames}")
            else:
                # ควรจะมี rename อย่างน้อย 1 อันเสมอ (คอลัมน์แรก -> id) ถ้าโหลดสำเร็จ
                self.log_message(f"Warning: No columns were renamed. Expected at least '{first_col_name}' -> '{self.id_col_target}'.")


            # --- Clean Data (ใช้ชื่อคอลัมน์หลัง Rename ถ้ามี, หรือชื่อเดิมถ้าไม่มี) ---
            # ใช้ชื่อคอลัมน์ *หลังจาก* rename ถ้ามีการ rename เกิดขึ้น
            age_col_to_clean = AGE_COLUMN_ORIGINAL # ถ้า AGE ไม่ได้ถูก rename ก็ใช้ชื่อเดิม
            # if AGE_COLUMN_ORIGINAL in actual_renames: age_col_to_clean = actual_renames[AGE_COLUMN_ORIGINAL] # ถ้ามีการ rename AGE

            if age_col_to_clean in temp_df.columns:
                temp_df[age_col_to_clean] = temp_df[age_col_to_clean].astype(str).str.strip()
                temp_df[age_col_to_clean] = temp_df[age_col_to_clean].replace({
                    '18 – 34 ปี': '18-34', '35 – 49 ปี': '35-49', '50 – 60 ปี': '50-60',
                    '18 - 34 ปี': '18-34', '35 - 49 ปี': '35-49', '50 - 60 ปี': '50-60'
                })
                self.log_message(f"Cleaned values in column '{age_col_to_clean}'.")
            else:
                self.log_message(f"Info: Column '{age_col_to_clean}' for age cleaning not found.")

            # ใช้ชื่อเป้าหมาย (ses_col_target) ถ้า SES ถูก rename สำเร็จ, มิฉะนั้นใช้ชื่อเดิม
            ses_col_to_clean = ses_col_target if SES_COLUMN_ORIGINAL in actual_renames else SES_COLUMN_ORIGINAL
            if ses_col_to_clean in temp_df.columns:
                temp_df[ses_col_to_clean] = temp_df[ses_col_to_clean].astype(str).str.strip()
                temp_df[ses_col_to_clean] = temp_df[ses_col_to_clean].replace({
                    'SES AB': 'AB', 'SES CD': 'CD'
                })
                self.log_message(f"Cleaned SES values in column '{ses_col_to_clean}'.")
            # ตรวจสอบเพิ่มเติมเผื่อกรณีที่ rename ไม่สำเร็จแต่ชื่อเดิมมีอยู่
            elif SES_COLUMN_ORIGINAL in temp_df.columns and SES_COLUMN_ORIGINAL != ses_col_to_clean:
                self.log_message(f"Info: Renamed SES column '{ses_col_to_clean}' not found, but original '{SES_COLUMN_ORIGINAL}' exists. Cleaning might be skipped or need adjustment.")
            else:
                self.log_message(f"Info: Column '{ses_col_to_clean}' (or original '{SES_COLUMN_ORIGINAL}') for SES cleaning not found.")


            # --- Final Check and Store ---
            # ตรวจสอบว่าการ rename คอลัมน์แรกเป็น 'id' สำเร็จหรือไม่
            if self.id_col_target not in temp_df.columns:
                raise KeyError(f"Mandatory ID column '{self.id_col_target}' was not created. Failed to rename the first column ('{first_col_name}') correctly.")

            self.cleaned_df = temp_df
            self.file_path.set(os.path.basename(filepath))
            available_cols = list(self.cleaned_df.columns)
            self.log_message(f"Dataset processed. {len(self.cleaned_df)} records available.")
            self.log_message(f"Available columns listed. Select desired dimension(s) in each Quota Set tab.")

            # --- Populate Dimension Listboxes (เหมือนเดิม) ---
            dim_options = [col for col in available_cols if col != self.id_col_target] # เอาคอลัมน์ 'id' ออก
            for i in range(self.num_quota_sets):
                listbox = self.quota_dim_listboxes[i]
                if listbox:
                    try:
                        listbox.config(state=tk.NORMAL)
                        listbox.delete(0, tk.END)
                        for option in dim_options:
                            listbox.insert(tk.END, option)
                        current_state = tk.NORMAL if self.quota_enable_vars[i].get() or i == 0 else tk.DISABLED
                        listbox.config(state=current_state)
                    except tk.TclError: pass

            self.btn_run['state'] = tk.NORMAL
            self.sampling_results = None

        # --- Exception Handling (ปรับปรุงเล็กน้อย) ---
        except FileNotFoundError: messagebox.showerror("Error", f"File not found: {filepath}"); self.file_path.set("Error loading file."); self.loaded_df = None; self.cleaned_df = None; self.btn_run['state'] = tk.DISABLED
        except KeyError as ke:
            # ทำให้ข้อความ Error ชัดเจนขึ้นว่าอาจเกิดจาก Rename ไม่สำเร็จ
            messagebox.showerror("Error Loading", f"Column processing error: {ke}.\nThis might indicate an issue renaming the first column or finding other specified columns (e.g., SES).");
            self.log_message(f"Error: {ke}"); self.file_path.set("Column error."); self.loaded_df = None; self.cleaned_df = None; self.btn_run['state'] = tk.DISABLED
        except ValueError as ve:
            # ดักจับ Error จากการแปลง type หรือ sheet ว่าง
            messagebox.showerror("Data Error", f"Error processing data: {ve}")
            self.log_message(f"Error: {ve}"); self.file_path.set("Data processing error."); self.loaded_df = None; self.cleaned_df = None; self.btn_run['state'] = tk.DISABLED
        except ImportError as ie:
            messagebox.showerror("Missing Library", f"Required library for reading Excel files is missing: {ie}\nPlease install it (e.g., `pip install openpyxl`) and restart.")
            self.log_message(f"Import Error: {ie}. Please install required Excel library (e.g., openpyxl).")
            self.file_path.set("Missing Excel library."); self.loaded_df = None; self.cleaned_df = None; self.btn_run['state'] = tk.DISABLED
        except Exception as e:
            messagebox.showerror("Error Loading", f"Unexpected error processing the file: {e}")
            self.log_message(f"Error: {e}"); self.file_path.set("Processing error."); self.loaded_df = None; self.cleaned_df = None; self.btn_run['state'] = tk.DISABLED

    # --- parse_dimensions (ไม่ต้องแก้ไข) ---
    def parse_dimensions(self, dim_string, quota_set_name="Quota"):
        """Parses comma-separated dimensions, strips whitespace, and validates against loaded data."""
        if not dim_string:
            messagebox.showerror("Input Error", f"{quota_set_name} dimensions cannot be empty.")
            return None
        dims = [d.strip() for d in dim_string.split(',') if d.strip()]
        if not dims:
            messagebox.showerror("Input Error", f"{quota_set_name} dimensions entry resulted in no valid dimension names.")
            return None

        if self.cleaned_df is None:
            messagebox.showerror("Error", "Dataset not loaded yet. Cannot validate dimensions.")
            return None

        available_cols = list(self.cleaned_df.columns)
        missing = [d for d in dims if d not in available_cols]
        if missing:
            messagebox.showerror("Dimension Error",
                                 f"{quota_set_name} dimension(s) not found in dataset:\n{', '.join(missing)}\n\n"
                                 f"Available columns are:\n{', '.join(available_cols)}")
            return None
        return dims

    # --- clear_dimension_values (ไม่ต้องแก้ไข) ---
    def clear_dimension_values(self, index):
        """Clears the comboboxes and their frame for a given quota set."""
        if index < 0 or index >= self.num_quota_sets: return
        self.quota_dim_comboboxes[index] = []
        frame_to_clear = self.quota_dim_value_frames[index]
        if frame_to_clear and frame_to_clear.winfo_exists():
            for widget in frame_to_clear.winfo_children():
                widget.destroy()

    # --- parse_quota_targets (ไม่ต้องแก้ไข) ---
    def parse_quota_targets(self, text_content, quota_set_name="Quota"):
        targets = {};
        if not text_content.strip():
             if "Set 1" in quota_set_name:
                 messagebox.showerror("Quota Error", f"{quota_set_name} targets cannot be empty."); return None
             else:
                 print(f"Info: {quota_set_name} targets text box is empty. Assuming no targets for this set.")
                 return {}

        lines = text_content.strip().split('\n');
        for i, line in enumerate(lines):
            line = line.strip();
            if not line or line.startswith('#'): continue
            try:
                parts = line.rsplit(':', 1);
                if len(parts) != 2: raise ValueError("Incorrect format (missing ':' or extra ':')");
                key_str, count_str = parts[0].strip(), parts[1].strip();
                key = ast.literal_eval(key_str);
                if not isinstance(key, tuple): raise TypeError("Key part is not a tuple");
                string_key = tuple(map(str, key))
                count = int(count_str);
                if count < 0: raise ValueError("Count cannot be negative")
                targets[string_key] = count
            except (ValueError, SyntaxError, TypeError) as e: messagebox.showerror("Parsing Error", f"Err parsing {quota_set_name} line {i+1}:\n'{line}'\nErr: {e}\nUse format ('V1','V2'):Count"); return None
            except Exception as e: messagebox.showerror("Parsing Error", f"Unexpected Err parsing {quota_set_name} line {i+1}:\n'{line}'\nErr: {e}"); return None
        return targets

    # --- load_dimension_values (ไม่ต้องแก้ไข) ---
    def load_dimension_values(self, index):
        """Loads unique values into comboboxes for the dimensions SELECTED in the listbox."""
        if self.cleaned_df is None:
            messagebox.showerror("Error", "Please load a dataset first.")
            return
        if index < 0 or index >= self.num_quota_sets:
            self.log_message(f"Error: Invalid quota set index {index}.")
            return

        quota_set_name = f"Quota Set {index+1}"
        dim_listbox = self.quota_dim_listboxes[index]
        if not dim_listbox:
             messagebox.showerror("Error", f"Dimension listbox not found for {quota_set_name}.")
             return

        selected_indices = dim_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Input Error", f"Please select at least one dimension from the list in {quota_set_name} before loading values.", parent=self.master)
            self.clear_dimension_values(index)
            return
        dims = [dim_listbox.get(i) for i in selected_indices]

        missing = [d for d in dims if d not in self.cleaned_df.columns]
        if missing:
             messagebox.showerror("Error", f"Selected dimension(s) not found in data: {missing}", parent=self.master)
             self.clear_dimension_values(index)
             return

        if self.quota_targets_data[index]:
            if messagebox.askyesno("Confirm Clear Targets",
                                  f"Loading values for dimensions [{', '.join(dims)}] in {quota_set_name} will CLEAR existing target cells defined for this set.\n\nDo you want to proceed?",
                                  parent=self.master):
                self.quota_targets_data[index] = {}
                self.update_target_display(index)
                self.log_message(f"Cleared previous targets for {quota_set_name}.")
            else:
                self.log_message(f"Loading dimension values cancelled by user for {quota_set_name}.")
                return

        self.clear_dimension_values(index)

        frame_to_populate = self.quota_dim_value_frames[index]
        if not frame_to_populate or not frame_to_populate.winfo_exists():
            self.log_message(f"Error: Value frame for {quota_set_name} not found or destroyed.")
            return

        new_combobox_list = []
        all_dims_valid = True
        try:
            for dim in dims:
                try:
                    # Convert to string before finding unique to handle mixed types better
                    unique_values_series = self.cleaned_df[dim].astype(str).dropna().unique()
                    # Filter out common representations of missing/empty after conversion
                    unique_values = sorted([val for val in unique_values_series if val and val.lower() not in ['nan', 'none', '<na>']])
                except Exception as e_val:
                     messagebox.showerror("Error", f"Error getting unique values for dimension '{dim}' in {quota_set_name}:\n{e_val}", parent=self.master)
                     all_dims_valid = False
                     break

                if not unique_values:
                     self.log_message(f"Warning: Dimension '{dim}' in {quota_set_name} has no valid, non-missing values.")
                     lbl_empty = ttk.Label(frame_to_populate, text=f"{dim}: (No values found)")
                     lbl_empty.pack(side=tk.LEFT, padx=5, pady=2)
                     new_combobox_list.append(None) # Placeholder for this dimension
                     continue

                lbl = ttk.Label(frame_to_populate, text=f"{dim}:")
                lbl.pack(side=tk.LEFT, padx=(10, 2), pady=2)
                combo = ttk.Combobox(frame_to_populate, values=unique_values, state='readonly', width=20)
                combo.pack(side=tk.LEFT, padx=(0, 10), pady=2)
                combo.current(0) # Select first value by default
                new_combobox_list.append(combo)

            if all_dims_valid:
                 self.quota_dim_comboboxes[index] = new_combobox_list
                 self.log_message(f"Loaded value selectors for dimensions [{', '.join(dims)}] in {quota_set_name}.")
            else:
                 self.clear_dimension_values(index)
                 self.log_message(f"Failed to load value selectors due to errors in {quota_set_name}.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to create dimension value selectors for {quota_set_name}: {e}", parent=self.master)
            self.clear_dimension_values(index)

    # --- add_target_cell (ไม่ต้องแก้ไข) ---
    def add_target_cell(self, index):
        """Adds a target cell based on combobox selections and count entry."""
        if index < 0 or index >= self.num_quota_sets: return

        quota_set_name = f"Quota Set {index+1}"
        comboboxes = self.quota_dim_comboboxes[index]
        count_entry = self.quota_target_count_entries[index]
        dim_listbox = self.quota_dim_listboxes[index]

        if comboboxes is None or not dim_listbox or count_entry is None:
             messagebox.showerror("Error", f"Components missing for {quota_set_name}. Please select dimensions and click 'Load Values' first.", parent=self.master)
             return

        selected_indices = dim_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", f"No dimensions selected in the list for {quota_set_name}.", parent=self.master)
            return
        dims = [dim_listbox.get(i) for i in selected_indices]

        actual_combobox_count = sum(1 for combo in comboboxes if combo is not None)
        if actual_combobox_count != len(dims):
             messagebox.showerror("Error", f"Mismatch between value selectors ({actual_combobox_count}) and selected dimensions ({len(dims)}) for {quota_set_name}. Try clicking 'Load Values' again.", parent=self.master)
             return

        selected_values = []
        valid_selection = True
        for i, combo in enumerate(comboboxes):
             # Skip if combobox was a None placeholder (dim had no values)
             if combo is None:
                  messagebox.showerror("Internal Error", f"Cannot add target: Dimension '{dims[i]}' had no values to select from.", parent=self.master)
                  valid_selection = False
                  break
             value = combo.get()
             if not value:
                  messagebox.showerror("Input Error", f"Please select a value for dimension '{dims[i]}' in {quota_set_name}.", parent=self.master)
                  valid_selection = False
                  break
             selected_values.append(value) # Values are already strings

        if not valid_selection: return

        target_key = tuple(selected_values)

        count_str = count_entry.get()
        try:
            if not count_str: raise ValueError("Target count cannot be empty.")
            count = int(count_str)
            if count < 0:
                 messagebox.showerror("Input Error", f"Target count must be zero or positive in {quota_set_name}.", parent=self.master)
                 return
        except ValueError:
            messagebox.showerror("Input Error", f"Invalid target count '{count_str}'. Enter a whole number (0+) in {quota_set_name}.", parent=self.master)
            return

        action = "Updated" if target_key in self.quota_targets_data[index] else "Added"
        self.quota_targets_data[index][target_key] = count
        self.log_message(f"{action} Target for {quota_set_name}: {target_key}: {count}")
        self.update_target_display(index)
        count_entry.delete(0, tk.END)


    # --- remove_target_cell (ไม่ต้องแก้ไข) ---
    def remove_target_cell(self, index):
        """Removes the selected target cell(s) from the listbox and internal data."""
        if index < 0 or index >= self.num_quota_sets: return
        listbox = self.quota_target_listboxes[index]
        if not listbox or not listbox.winfo_exists():
             self.log_message(f"Error: Target listbox for Quota Set {index+1} not found.")
             return

        selected_indices = listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("Info", "Please select target cell(s) from the list to remove.", parent=self.master)
            return

        quota_set_name = f"Quota Set {index+1}"
        items_to_remove = [listbox.get(i) for i in selected_indices]
        removed_count = 0
        parse_errors = 0

        for item_str in items_to_remove:
            try:
                key_str = item_str.rsplit(':', 1)[0].strip()
                target_key = ast.literal_eval(key_str)
                if not isinstance(target_key, tuple): raise ValueError("Parsed key is not a tuple")

                if target_key in self.quota_targets_data[index]:
                    del self.quota_targets_data[index][target_key]
                    removed_count += 1
                else:
                    self.log_message(f"Warning: Key {target_key} from '{item_str}' not found in data for {quota_set_name}.")

            except (IndexError, SyntaxError, ValueError, TypeError) as e:
                self.log_message(f"Error parsing item to remove: '{item_str}'. Error: {e}")
                parse_errors += 1
            except Exception as e:
                 self.log_message(f"Unexpected error removing item '{item_str}': {e}")
                 parse_errors += 1

        if parse_errors > 0:
             messagebox.showwarning("Warning", f"Could not parse {parse_errors} selected item(s). They were not removed.", parent=self.master)

        if removed_count > 0:
             self.log_message(f"Removed {removed_count} target cell(s) from {quota_set_name}.")
             self.update_target_display(index)
        elif parse_errors > 0 and removed_count == 0:
             self.log_message(f"No targets removed due to parsing errors for {quota_set_name}.")


    # --- update_target_display (ไม่ต้องแก้ไข) ---
    def update_target_display(self, index):
        """Clears and repopulates the target listbox for the given quota set."""
        if index < 0 or index >= self.num_quota_sets: return
        listbox = self.quota_target_listboxes[index]
        if not listbox or not listbox.winfo_exists(): return

        try:
             listbox.config(state=tk.NORMAL)
             listbox.delete(0, tk.END)
             sorted_items = sorted(self.quota_targets_data[index].items())
             for key, count in sorted_items:
                 listbox.insert(tk.END, f"{key}: {count}")
        except tk.TclError: pass # Widget might be destroyed
        except Exception as e:
             self.log_message(f"Error updating target display for Q{index+1}: {e}")


    # --- run_sampling (ไม่ต้องแก้ไข) ---
    def run_sampling(self):
        if self.cleaned_df is None: messagebox.showerror("Error", "Load dataset first."); return
        self.clear_log(); self.log_message("Starting sampling process...");
        self.btn_run['state'] = tk.DISABLED; self.btn_export['state'] = tk.DISABLED; self.master.update()

        quota_definitions = []
        parsed_q1_targets = None

        try:
            for i in range(self.num_quota_sets):
                if i == 0 or self.quota_enable_vars[i].get():
                    quota_set_name = f"Quota Set {i+1}"
                    dim_listbox = self.quota_dim_listboxes[i]

                    if not dim_listbox:
                         if i == 0: raise ValueError(f"Dimension listbox missing for required {quota_set_name}")
                         else: self.log_message(f"Warning: Skipping {quota_set_name}, dimension listbox not found."); continue

                    selected_indices = dim_listbox.curselection()
                    if not selected_indices:
                        if i == 0: raise ValueError(f"{quota_set_name} requires at least one dimension to be selected.")
                        elif self.quota_targets_data[i]: raise ValueError(f"{quota_set_name} is enabled and has targets, but no dimensions are selected.")
                        else: self.log_message(f"Info: Skipping enabled {quota_set_name} (no dimensions selected)."); continue

                    q_dims = [dim_listbox.get(k) for k in selected_indices]
                    q_targets = self.quota_targets_data[i].copy()

                    if i == 0 and not q_targets:
                         raise ValueError(f"{quota_set_name} requires at least one target cell.")
                    if i > 0 and self.quota_enable_vars[i].get() and not q_targets:
                         self.log_message(f"Info: {quota_set_name} is enabled but has no target cells defined.")

                    quota_definitions.append((q_dims, q_targets))
                    self.log_message(f"Using {quota_set_name}: Dims={q_dims}, Targets={len(q_targets)} cells.")
                    if i == 0: parsed_q1_targets = q_targets

            if not quota_definitions: raise ValueError("No valid quota definitions. Quota Set 1 with selected dimensions and targets is required.")
            if parsed_q1_targets is None: raise ValueError("Quota Set 1 definition is required but was not processed.")

            # Check target sum mismatches
            q1_sum = sum(parsed_q1_targets.values()) if parsed_q1_targets else 0
            active_quota_indices = [idx for idx in range(self.num_quota_sets) if idx == 0 or self.quota_enable_vars[idx].get()]
            processed_idx_map = {proc_idx: active_idx for proc_idx, active_idx in enumerate(active_quota_indices)}

            for proc_idx, (dims, targets) in enumerate(quota_definitions):
                 if proc_idx == 0: continue # Skip Q1 itself
                 current_sum = sum(targets.values()) if targets else 0
                 original_set_num = processed_idx_map.get(proc_idx, -1) + 1 # Get original set number
                 if targets and current_sum != q1_sum:
                     self.log_message(f"Warning: Target sum mismatch! Quota Set 1 ({q1_sum}) vs Set {original_set_num} ({current_sum}). Sampling based on Q1 target size.")

            self.log_message(f"\nRunning sampling with {len(quota_definitions)} ACTIVE quota set(s)...")
            results = flexible_quota_sampling(self.cleaned_df, quota_definitions, id_col=self.id_col_target)

            # Process and Display Results
            self.sampling_results = results; selected_ids, final_counts_list, unmet_list = results
            target_sample_size = q1_sum
            self.log_message(f"\n--- Sampling Finished ---\nSelected: {len(selected_ids)} (Target based on Q1: {target_sample_size})")

            for proc_idx, final_counts in enumerate(final_counts_list):
                 original_set_num = processed_idx_map.get(proc_idx, -1) + 1
                 original_dims, original_targets = quota_definitions[proc_idx]
                 quota_name = f"Quota Set {original_set_num}"
                 self.log_message(f"\n--- {quota_name} Summary (Selected vs Target) ---")
                 if not original_targets: self.log_message("- No targets were defined for this set."); continue

                 all_keys_summary = set(final_counts.keys()) | set(original_targets.keys())
                 if not all_keys_summary: self.log_message("- No targets defined or selected."); continue

                 for key in sorted(list(all_keys_summary)):
                     sel_count = final_counts.get(key, 0)
                     tar_count = original_targets.get(key, 0)
                     self.log_message(f"- {key}: Selected={sel_count} (Target={tar_count})")

                 current_unmet = unmet_list[proc_idx]
                 if current_unmet:
                     self.log_message(f"\n{quota_name} - Unmet Targets:")
                     for k, sf in sorted(current_unmet.items()):
                         t = original_targets.get(k, 0)
                         s = final_counts.get(k, 0)
                         self.log_message(f"  - {k}: Shortfall={sf} (Target={t}, Selected={s})")
                 else:
                     self.log_message(f"{quota_name}: All defined targets met or exceeded.")


            self.btn_export['state'] = tk.NORMAL; self.log_message("\nตัด Quota complete. กด Export Excelได้เลย!!!.")
            # +++++ เพิ่มบรรทัดนี้ +++++
            messagebox.showinfo("สำเร็จ", "ตัดชุดเรียบร้อย\nกด Export Excel ได้เลย!!", parent=self.master)

        except ValueError as ve: messagebox.showerror("Input Error", f"{ve}", parent=self.master); self.log_message(f"ERROR: {ve}")
        except Exception as e: messagebox.showerror("Sampling Error", f"An unexpected error during sampling: {e}", parent=self.master); self.log_message(f"ERROR: {e}")
        finally: self.btn_run['state'] = tk.NORMAL


    # --- export_results (ไม่ต้องแก้ไข - ยังคง export เป็น .xlsx) ---
    def export_results(self):
        if self.sampling_results is None or self.cleaned_df is None: messagebox.showerror("Error", "No sampling results to export."); return
        filepath = filedialog.asksaveasfilename(title="Save Results", defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"), ("All", "*.*")))
        if not filepath: return
        self.log_message(f"Exporting to: {filepath}..."); self.btn_export['state'] = tk.DISABLED; self.master.update()

        try:
            selected_ids, final_counts_list, unmet_list = self.sampling_results
            selected_population_df = pd.DataFrame()
            if selected_ids:
                 try:
                    df_to_filter = self.cleaned_df.copy()
                    df_to_filter[self.id_col_target] = df_to_filter[self.id_col_target].astype(str)
                    selected_population_df = df_to_filter[df_to_filter[self.id_col_target].isin(selected_ids)].copy()
                    if selected_population_df.empty and selected_ids: self.log_message("ERROR Export: Filtering resulted in empty DataFrame!")
                 except Exception as e_conv: self.log_message(f"ERROR Export: Failed to convert/filter DataFrame ID: {e_conv}")
            else: self.log_message("DEBUG Export: No selected IDs to filter.")

            # Get ACTIVE definitions for export sheets
            quota_definitions_export = []
            active_quota_indices = [] # Track original indices
            try:
                for i in range(self.num_quota_sets):
                    if i == 0 or self.quota_enable_vars[i].get():
                        dim_listbox = self.quota_dim_listboxes[i]
                        if not dim_listbox:
                             if i==0: raise ValueError("Export Error: Dim listbox missing for Q1.")
                             else: self.log_message(f"Warn Export: Dim listbox missing for Q{i+1}, skipping summary."); continue

                        selected_indices = dim_listbox.curselection()
                        if not selected_indices:
                             if i == 0: raise ValueError("Export Error: Q1 requires selected dims.")
                             elif self.quota_targets_data[i]: raise ValueError(f"Export Error: Q{i+1} has targets but no selected dims.")
                             else: self.log_message(f"Info Export: Skipping summary for Q{i+1}, no dims selected."); continue

                        q_dims = [dim_listbox.get(k) for k in selected_indices]
                        q_targets = self.quota_targets_data[i].copy()

                        if i == 0 and not q_targets: raise ValueError("Export Error: Q1 requires targets for summary.")
                        elif i > 0 and self.quota_enable_vars[i].get() and not q_targets: self.log_message(f"Info Export: Q{i+1} included with 0 targets.")

                        quota_definitions_export.append({'dims': q_dims, 'targets': q_targets})
                        active_quota_indices.append(i) # Store the original index

            except ValueError as ve: self.log_message(f"ERROR Export Prep: {ve}"); raise ve
            except Exception as parse_err: self.log_message(f"ERROR Export Prep: {parse_err}"); raise parse_err

            if not quota_definitions_export: self.log_message("Warning Export: No active quotas for summary sheets.")

            # Write Excel
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                # Sheet: Selected Individuals
                try:
                    if not selected_population_df.empty:
                        self.log_message("DEBUG Export: Writing 'Selected_Individuals'...")
                        selected_population_df.to_excel(writer, sheet_name='Selected_Individuals', index=False)
                    elif selected_ids: self.log_message("INFO Export: 'Selected_Individuals' skipped (filtering failed).")
                    else: self.log_message("INFO Export: 'Selected_Individuals' skipped (no IDs selected).")
                except Exception as e_sheet1: self.log_message(f"ERROR Exporting Sheet 'Selected_Individuals': {e_sheet1}")

                # Sheets for Quota Summaries & Unmet
                max_summary_index = min(len(final_counts_list), len(quota_definitions_export))
                unmet_df_list = []

                for i in range(max_summary_index):
                    final_counts = final_counts_list[i]
                    export_def = quota_definitions_export[i]
                    original_set_index = active_quota_indices[i] # Get original index
                    sheet_name = f'Quota{original_set_index+1}_Summary'
                    current_targets = export_def['targets']
                    dims = export_def['dims']
                    self.log_message(f"DEBUG Export: Preparing '{sheet_name}'...")

                    try:
                         key_len = len(dims)
                         result_list_for_sheet = []
                         all_keys_for_sheet = set(final_counts.keys()) | set(current_targets.keys())
                         for key in sorted(list(all_keys_for_sheet)):
                              if isinstance(key, tuple) and len(key) == key_len:
                                   row_data = {'Dim{}'.format(j+1): key[j] for j in range(key_len)}
                                   row_data['Selected'] = final_counts.get(key, 0)
                                   row_data['Target'] = current_targets.get(key, 0)
                                   result_list_for_sheet.append(row_data)
                              else: self.log_message(f"WARN Export: Invalid key {key} in {sheet_name}. Skipping.")
                         if result_list_for_sheet:
                              summary_df = pd.DataFrame(result_list_for_sheet)
                              dim_rename_map = {f'Dim{j+1}': dims[j] for j in range(key_len)}
                              summary_df.rename(columns=dim_rename_map, inplace=True)
                              # Pivot logic for 2 dimensions
                              if key_len == 2 and len(summary_df) > 0:
                                   dim1_name, dim2_name = dims[0], dims[1]
                                   try:
                                        result_pivot = summary_df.pivot_table(index=dim1_name, columns=dim2_name, values='Selected', fill_value=0, aggfunc='sum')
                                        target_pivot = summary_df.pivot_table(index=dim1_name, columns=dim2_name, values='Target', fill_value=0, aggfunc='sum')
                                        workbook = writer.book
                                        worksheet = writer.sheets.get(sheet_name)
                                        if worksheet is None:
                                             # If sheet doesn't exist yet (first pivot write)
                                             # Write target first
                                             target_pivot.to_excel(writer, sheet_name=sheet_name, startrow=1, startcol=0)
                                             worksheet = writer.sheets[sheet_name] # Get the created sheet
                                             worksheet.write(0, 0, "Target Counts:")
                                             # Write selected below target
                                             start_row_selected = target_pivot.shape[0] + 3
                                             worksheet.write(start_row_selected -1 , 0, 'Selected Counts:')
                                             result_pivot.to_excel(writer, sheet_name=sheet_name, startrow=start_row_selected, startcol=0)
                                        else: # Sheet exists, append selected pivot
                                             start_row_selected = target_pivot.shape[0] + 3
                                             worksheet.write(start_row_selected -1 , 0, 'Selected Counts:')
                                             result_pivot.to_excel(writer, sheet_name=sheet_name, startrow=start_row_selected, startcol=0)

                                        self.log_message(f"DEBUG Export: '{sheet_name}' pivot table write finished.")
                                   except Exception as e_pivot:
                                        self.log_message(f"WARN Export: Pivot failed for '{sheet_name}': {e_pivot}. Writing list.")
                                        summary_df.to_excel(writer, sheet_name=sheet_name, index=False) # Fallback to list
                              else: # 1 or 3+ dimensions, or pivot failed
                                   summary_df.to_excel(writer, sheet_name=sheet_name, index=False)
                                   self.log_message(f"DEBUG Export: '{sheet_name}' list write finished.")
                         else: self.log_message(f"INFO Export: No data for '{sheet_name}'.")
                    except Exception as e_sheet_sum: self.log_message(f"ERROR Exporting {sheet_name}: {e_sheet_sum}")


                    # Collect Unmet Data
                    if i < len(unmet_list): # Ensure unmet_list has entry for this index
                         unmet_dict = unmet_list[i]
                         if unmet_dict:
                              quota_name_unmet = f"QuotaSet_{original_set_index+1}"
                              current_counts = final_counts_list[i]
                              key_len = len(dims)
                              try:
                                  unmet_rows = []
                                  for key, sf in unmet_dict.items():
                                       if isinstance(key, tuple) and len(key) == key_len:
                                           row_data = {'QuotaSet': quota_name_unmet}
                                           row_data.update({f'Dim{j+1}_Name': dims[j] for j in range(key_len)}) # Add Dim Names
                                           row_data.update({f'Dim{j+1}_Value': key[j] for j in range(key_len)}) # Add Dim Values
                                           row_data['Target'] = current_targets.get(key, 0)
                                           row_data['Selected'] = current_counts.get(key, 0)
                                           row_data['Shortfall'] = sf
                                           unmet_rows.append(row_data)
                                       else: self.log_message(f"WARN Export Unmet: Invalid key {key}. Skipping.")
                                  if unmet_rows:
                                       unmet_df_list.append(pd.DataFrame(unmet_rows))
                              except Exception as e_unmet_prep: self.log_message(f"ERROR Prep Unmet Q{original_set_index+1}: {e_unmet_prep}")
                    else:
                          self.log_message(f"Warn Export: Missing entry in unmet_list for index {i} (Quota Set {original_set_index+1})")


                # Write Unmet Quotas Sheet (after loop)
                if unmet_df_list:
                     try:
                        full_unmet_df = pd.concat(unmet_df_list, ignore_index=True)
                        if not full_unmet_df.empty:
                            # Reorder columns for clarity
                            dim_cols = [f'Dim{j+1}_Name' for j in range(key_len)] + [f'Dim{j+1}_Value' for j in range(key_len)]
                            other_cols = ['QuotaSet', 'Target', 'Selected', 'Shortfall']
                            # Handle cases where key_len might vary if different sets failed? Unlikely with current structure.
                            # Use existing columns if some DimX cols are missing
                            final_cols_order = other_cols[:1] + [c for c in dim_cols if c in full_unmet_df.columns] + other_cols[1:]
                            full_unmet_df = full_unmet_df[final_cols_order]

                            self.log_message("DEBUG Export: Writing 'Unmet_Quotas' sheet...")
                            full_unmet_df.to_excel(writer, sheet_name='Unmet_Quotas', index=False)
                        else: self.log_message("INFO Export: 'Unmet_Quotas' sheet skipped (no data after concat).")
                     except Exception as e_unmet_write: self.log_message(f"ERROR Exporting Unmet Sheet: {e_unmet_write}")
                else: self.log_message("INFO Export: No unmet quotas found.")


            self.log_message(f"Excel export finished for {filepath}"); messagebox.showinfo("Export OK", f"Saved to:\n{filepath}", parent=self.master)
        except ValueError as ve:
             messagebox.showerror("Export Error", f"Cannot export summaries:\n{ve}", parent=self.master); self.log_message(f"ERROR exporting: {ve}")
        except Exception as e: messagebox.showerror("Export Error", f"An unexpected error during export: {e}", parent=self.master); self.log_message(f"ERROR exporting: {e}")
        finally: self.btn_export['state'] = tk.NORMAL


# <<< START OF CHANGES >>>
# --- ฟังก์ชัน Entry Point ใหม่ (สำหรับให้ Launcher เรียก) ---
def run_this_app(working_dir=None): # ชื่อฟังก์ชันนี้จะถูกใช้ใน Launcher
    """
    ฟังก์ชันหลักสำหรับสร้างและรัน QuotaSamplerApp.
    """
    print(f"--- QUOTA_SAMPLER_INFO: Starting 'QuotaSamplerApp' via run_this_app() ---")
    try:
        # --- โค้ดที่ย้ายมาจาก if __name__ == "__main__": เดิมจะมาอยู่ที่นี่ ---
        root = tk.Tk()
        app = QuotaSamplerApp(root) # สร้าง Instance ของ GUI App
        root.mainloop() # เริ่มการทำงานของหน้าต่าง GUI

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