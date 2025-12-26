import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import pyreadstat # For reading SPSS files
from datetime import datetime
import openpyxl # For Excel export
import webbrowser
import tkinter as tk
import sys
import threading # For background tasks

# --- Imports for Google Sheets ---
import gspread
from google.oauth2.service_account import Credentials # Recommended for google-auth
# ---------------------------------


# Set default theme for customtkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class SpssExporterApp(ctk.CTk):
    MAX_DISPLAY_ITEMS_IN_DIALOG = 100 # Max items to display in dialog lists

    def _center_window(self, window_to_center, width, height):
        """Centers the given window on the screen."""
        window_to_center.update_idletasks()
        screen_width = window_to_center.winfo_screenwidth()
        screen_height = window_to_center.winfo_screenheight()
        x_coordinate = int((screen_width / 2) - (width / 2))
        y_coordinate = int((screen_height / 2) - (height / 2))
        window_to_center.geometry(f"{width}x{height}+{x_coordinate}+{y_coordinate}")

    def __init__(self):
        super().__init__()

        self.title("เก็บข้อมูล Norm For DP")

        app_width = 1000
        app_height = 700
        self._center_window(self, app_width, app_height)

        self.label_font_bold = ctk.CTkFont(weight="bold")
        self.panel_label_font_bold = ctk.CTkFont(size=14, weight="bold")
        self.dialog_info_font = ctk.CTkFont(size=10)


        self.spss_df_codes = None
        self.spss_df_labels = None
        self.spss_meta = None
        self.spss_variables = []

        # --- Google Sheet URL and Settings ---
        self.TARGET_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1CU0KjP3eAYRuJk7uz5rBR0GQ8uDUH4zuYnr6rnSKMwU/edit?gid=106958934#gid=106958934"
        self.SERVICE_ACCOUNT_FILE_PATH = 'Test3.json'  # <<<*** Specify Service Account Key File Path here ***
        self.INDEX_SHEET_NAME_FOR_COUNT = "Index"      # <<<*** Specify Sheet name for counting Jobs here ***
        self.UNIQUE_JOB_IDENTIFIER_COLUMN = "Project #" # <<<*** Column for unique Job identifier ***
        # --------------------------------------

        self.TARGET_Norm_URL = "https://script.google.com/macros/s/AKfycbxv0suQxHaDNttsaHpBIo2zqf-QY8IjaXKloqYoKmKSSbOUmwLeNUxZdH7-bxQr8dOz/exec"

        # --- Configuration for Fixed Inputs (Left Panel) ---
        self.current_year_str = str(datetime.now().year)
        year_implemented_options = [str(y) for y in range(int(self.current_year_str) - 4, int(self.current_year_str) + 5)]
        survey_method_options = ["CLT", "HUT", "Online survey", "D2D", "Intercept", "CATI", "Others"]
        category_options = [
            "Food","Snacks","Beverages","Alchoholic beverage","Household care",
            "Female Cosmetics and Fragrance","Male Cosmetics and Fragrance",
            "Facial care","Body & Haircare","Sanitary products","Diaper for baby",
            "Other baby care products","Diaper for adult","Pet care products",
            "Healthcare (OTC, functional foods & others)","Tobacco","Home appliances",
            "Banking","Telecom","Real Estate","Advertisement","Services Others","Others"
        ]
        evaluation_target_options = [
            "Concept","Product","Taste","Fragrance","Package","Brand","Ad/Campaign","Service"
        ]
        scale_options = ["3-point scale", "4-point scale", "5-point scale","7-point scale","9-point scale","10-point scale", "11-point scale"]
        self.username_placeholder = "" # Placeholder for username combobox
        username_options_with_placeholder = [self.username_placeholder] + ["Songklod","Kanokrat","Pornpan","Ittichai","Chanida"]


        self.fixed_input_fields = {
            "Username": {"type": "combobox", "options": username_options_with_placeholder, "widget": None, "var": None, "default": self.username_placeholder},
            "Timestamp": {"type": "entry", "widget": None, "var": None, "default": "Auto (on export)", "hidden": True},
            "Project #": {"type": "entry", "widget": None, "var": None, "default": ""},
            "Project Name": {"type": "entry", "widget": None, "var": None, "default": ""},
            "Year Implemented": {"type": "combobox", "options": year_implemented_options, "widget": None, "var": None, "default": self.current_year_str},
            "Survey Method": {"type": "multiselect_dialog_button", "options": survey_method_options, "widget_button": None, "widget_display": None, "var": None, "selected_values": [], "default_selected": []},
            "Category": {"type": "multiselect_dialog_button", "options": category_options, "widget_button": None, "widget_display": None, "var": None, "selected_values": [], "default_selected": []},
            "Evaluation Target": {"type": "multiselect_dialog_button", "options": evaluation_target_options, "widget_button": None, "widget_display": None, "var": None, "selected_values": [], "default_selected": []},
            "Scale": {"type": "combobox", "options": scale_options, "widget": None, "var": None, "default": "5-point scale"},
        }

        self.demographic_fields = ["Gender", "Age (NA)", "Age (Range)", "Area", "SES"]
        self.demographic_var_selections = {field: None for field in self.demographic_fields}
        self.demographic_widgets = {}

        self.channel_custom_names = [
            "01 PI No price", "02 PI Price", "03 OL", "04 Attractive",
            "05 Overall Appealing", "06 Usage Intention", "07 Newness",
            "08 Uniqueness", "09 Relevancy", "10 Credibility",
            "11 Concept Match", "12 Compare Current", "13 Overall Satisfaction",
            "14 Fit Brand", "15 Recommend"
        ]
        self.num_channels = len(self.channel_custom_names)
        self.channel_selections = {i: [] for i in range(self.num_channels)}
        self.channel_widgets = []

        # --- Variable for Total Jobs ---
        self.total_jobs_var = ctk.StringVar(value="Total Jobs: Loading...")
        # -----------------------------

        # --- Main UI Layout ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(1, weight=1)

        top_frame = ctk.CTkFrame(self)
        top_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=(10,5), sticky="ew")

        self.load_button = ctk.CTkButton(top_frame, text="Load SPSS File (.sav)", command=self.load_spss_file)
        self.load_button.pack(side="left", padx=5, pady=5)

        self.total_jobs_label = ctk.CTkLabel(
            top_frame,
            textvariable=self.total_jobs_var,
            font=self.label_font_bold,
            text_color="red"
        )
        self.total_jobs_label.pack(side="right", padx=(10, 5), pady=5)

        self.file_label = ctk.CTkLabel(top_frame, text="No file loaded.", anchor="w")
        self.file_label.pack(side="left", padx=5, pady=5, fill="x", expand=True)


        left_panel = ctk.CTkScrollableFrame(self,
                                            label_text="Project Info & Demographics",
                                            label_font=self.panel_label_font_bold)
        left_panel.grid(row=1, column=0, padx=(10, 5), pady=5, sticky="nsew")

        for field_name, config in self.fixed_input_fields.items():
            if config.get("hidden", False):
                config["var"] = ctk.StringVar(value=config.get("default", ""))
                continue
            item_frame = ctk.CTkFrame(left_panel)
            item_frame.pack(fill="x", padx=5, pady=(3, 4))
            label = ctk.CTkLabel(item_frame, text=field_name + ":", width=140, anchor="w", font=self.label_font_bold)
            label.pack(side="left", padx=(0,5))
            config["var"] = ctk.StringVar()
            default_value_to_set = config.get("default", "")
            if config["type"] == "entry":
                widget = ctk.CTkEntry(item_frame, textvariable=config["var"])
                config["var"].set(default_value_to_set)
                widget.pack(side="left", fill="x", expand=True)
                config["widget"] = widget
            elif config["type"] == "combobox":
                options = config.get("options", [])
                widget = ctk.CTkComboBox(item_frame, variable=config["var"], values=options, state="readonly")
                if default_value_to_set in options:
                    config["var"].set(default_value_to_set)
                elif options:
                    config["var"].set(options[0])
                else:
                    config["var"].set("")
                widget.pack(side="left", fill="x", expand=True)
                config["widget"] = widget
            elif config["type"] == "multiselect_dialog_button":
                config["selected_values"] = list(config.get("default_selected", []))
                config["var"].set(", ".join(config["selected_values"]))
                display_label = ctk.CTkEntry(item_frame, textvariable=config["var"], state="disabled", fg_color="gray80")
                display_label.pack(side="left", fill="x", expand=True, padx=(0,5))
                config["widget_display"] = display_label
                button = ctk.CTkButton(item_frame, text="Select...", width=80,
                                       command=lambda fn=field_name: self.open_multiselect_dialog(fn))
                button.pack(side="left")
                config["widget_button"] = button

        for field_name in self.demographic_fields:
            item_frame = ctk.CTkFrame(left_panel)
            item_frame.pack(fill="x", padx=5, pady=(3, 4))
            label = ctk.CTkLabel(item_frame, text=field_name + ":", width=140, anchor="w", font=self.label_font_bold)
            label.pack(side="left", padx=(0,5))
            display_spss_var_label = ctk.CTkLabel(item_frame, text="None selected", anchor="w")
            display_spss_var_label.pack(side="left", fill="x", expand=True, padx=5)
            select_spss_var_button = ctk.CTkButton(item_frame, text="Select Var", width=100,
                                                   command=lambda fn=field_name, dv_label=display_spss_var_label: self.open_demographic_var_selection_dialog(fn, dv_label))
            select_spss_var_button.pack(side="left")
            self.demographic_widgets[field_name] = {"display": display_spss_var_label, "button": select_spss_var_button}

        right_panel = ctk.CTkScrollableFrame(self,
                                             label_text="Concept Evaluation Channels (1-15)",
                                             label_font=self.panel_label_font_bold)
        right_panel.grid(row=1, column=1, padx=(5, 10), pady=5, sticky="nsew")

        for i in range(self.num_channels):
            channel_frame = ctk.CTkFrame(right_panel)
            channel_frame.pack(fill="x", padx=5, pady=(3,4))
            label_text = self.channel_custom_names[i] + ":"
            channel_label_widget = ctk.CTkLabel(channel_frame, text=label_text, anchor="w", width=170, font=self.label_font_bold)
            channel_label_widget.pack(side="left", padx=(0,5))
            selected_vars_display_textbox = ctk.CTkTextbox(channel_frame, height=40, state="disabled", wrap="word")
            selected_vars_display_textbox.pack(side="left", fill="x", expand=True, padx=5)
            select_channel_vars_button = ctk.CTkButton(channel_frame, text="Select Vars", width=100,
                                                       command=lambda ch_idx=i: self.open_channel_var_selection_dialog(ch_idx))
            select_channel_vars_button.pack(side="left")
            self.channel_widgets.append({"display": selected_vars_display_textbox})


        self.report_window_instance = None

        export_button_frame = ctk.CTkFrame(self)
        export_button_frame.grid(row=2, column=0, columnspan=2, pady=(5,10), sticky="ew")

        export_button_frame.grid_columnconfigure(0, weight=1)
        export_button_frame.grid_columnconfigure(1, weight=1)
        export_button_frame.grid_columnconfigure(2, weight=1)
        export_button_frame.grid_columnconfigure(3, weight=1)

        self.export_excel_button = ctk.CTkButton(export_button_frame, text="Export to Excel", command=self.export_to_excel, height=35)
        self.export_excel_button.grid(row=0, column=0, padx=(10,5), pady=10, sticky="ew")

        self.report_button = ctk.CTkButton(export_button_frame,
                                           text="Dashboard Report",
                                           command=self.open_Dashboard_link,
                                           height=35,
                                           fg_color="#00796B",
                                           hover_color="#004D40")
        self.report_button.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        self.open_gsheet_button = ctk.CTkButton(export_button_frame,
                                                text="คลิกเพื่อเปิด Google Sheet",
                                                command=self.open_google_sheet_link,
                                                height=35,
                                                fg_color="#F00A0A",
                                                hover_color="#EB4646")
        self.open_gsheet_button.grid(row=0, column=2, padx=5, pady=10, sticky="ew")

        self.gspread_button = ctk.CTkButton(export_button_frame,
                                            text="Save To Google Sheet",
                                            command=self.save_to_google_sheet,
                                            height=35,
                                            fg_color="darkgreen",
                                            hover_color="#179F05")
        self.gspread_button.grid(row=0, column=3, padx=(5,10), pady=10, sticky="ew")

        self._clear_all_selections_and_inputs()
        self.fetch_total_jobs_from_gsheet_threaded()

    def _clear_all_selections_and_inputs(self, clear_project_info=True, clear_demographics=True):
        """Clears all user inputs and selections to their default states.
        
        Args:
            clear_project_info (bool): Whether to clear project info fields (default: True)
            clear_demographics (bool): Whether to clear demographics selections (default: True)
        """
        # Clear fixed input fields only if clear_project_info is True
        if clear_project_info:
            for field_name, config in self.fixed_input_fields.items():
                if config.get("hidden", False):
                    if "var" in config and config["var"] is not None:
                        config["var"].set(config.get("default", ""))
                    continue

                if "var" in config and config["var"] is not None:
                    default_value = config.get("default", "")
                    if config["type"] == "combobox" and not default_value and config.get("options"):
                        default_value = config["options"][0]
                    elif config["type"] == "multiselect_dialog_button":
                        config["selected_values"] = list(config.get("default_selected", []))
                        default_value = ", ".join(config["selected_values"])
                    config["var"].set(default_value)

        # Clear demographics only if clear_demographics is True
        if clear_demographics:
            self.demographic_var_selections = {field: None for field in self.demographic_fields}
            for field_name in self.demographic_fields:
                if field_name in self.demographic_widgets and "display" in self.demographic_widgets[field_name]:
                    self.demographic_widgets[field_name]["display"].configure(text="None selected")

        # Always clear channel selections (concept evaluation channels)
        self.channel_selections = {i: [] for i in range(self.num_channels)}
        for i in range(self.num_channels):
            self.update_channel_display_text(i)
        
        if clear_project_info and clear_demographics:
            print("All selections and inputs cleared to default.")
        elif clear_project_info:
            print("Project info and channel selections cleared to default.")
        elif clear_demographics:
            print("Demographics and channel selections cleared to default.")
        else:
            print("Only channel selections cleared to default.")


    def open_google_sheet_link(self):
        try:
            if self.TARGET_GOOGLE_SHEET_URL:
                webbrowser.open_new_tab(self.TARGET_GOOGLE_SHEET_URL)
            else:
                messagebox.showwarning("No URL", "Target Google Sheet URL is not defined.")
        except Exception as e:
            messagebox.showerror("Error Opening Link", f"Could not open the Google Sheet link in browser:\n{str(e)}")

    def open_Dashboard_link(self):
        try:
            if self.TARGET_Norm_URL:
                webbrowser.open_new_tab(self.TARGET_Norm_URL)
            else:
                messagebox.showwarning("No URL", "Target Dashboard URL is not defined.")
        except Exception as e:
            messagebox.showerror("Error Opening Link", f"Could not open the Dashboard link in browser:\n{str(e)}")

    def load_spss_file(self):
        file_path = filedialog.askopenfilename(title="Select SPSS File", filetypes=(("SPSS files", "*.sav"), ("All files", "*.*")))
        if not file_path:
            return

        self.loading_dialog = ctk.CTkToplevel(self)
        self.loading_dialog.title("Loading SPSS")
        loading_dialog_width = 350
        loading_dialog_height = 120
        self._center_window(self.loading_dialog, loading_dialog_width, loading_dialog_height)
        self.loading_dialog.grab_set()
        self.loading_dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        ctk.CTkLabel(self.loading_dialog, text="กำลังโหลดไฟล์ SPSS กรุณารอสักครู่...\nขั้นตอนนี้อาจใช้เวลาหลายนาทีสำหรับไฟล์ขนาดใหญ่",
                     font=self.label_font_bold, justify="center").pack(pady=20, padx=20)
        self.loading_dialog.update()

        self.spss_load_results = {
            "df_codes": None, "meta": None, "success": False,
            "used_encoding": "N/A", "last_exception": None
        }

        def _load_file_in_thread():
            encodings_to_try = [None, 'utf-8', 'cp874', 'tis-620']

            for enc in encodings_to_try:
                try:
                    current_encoding_name = enc if enc else 'auto-detect by pyreadstat'
                    print(f"Attempting to load SPSS file with encoding: {current_encoding_name}")

                    temp_df_codes, temp_meta = pyreadstat.read_sav(
                        file_path,
                        apply_value_formats=False,
                        user_missing=False,
                        encoding=enc
                    )

                    if temp_df_codes is not None and not temp_df_codes.empty and temp_meta is not None:
                        self.spss_load_results["df_codes"] = temp_df_codes
                        self.spss_load_results["meta"] = temp_meta
                        self.spss_load_results["success"] = True
                        self.spss_load_results["used_encoding"] = current_encoding_name
                        print(f"Successfully loaded codes and meta with encoding: {current_encoding_name}")
                        break
                    else:
                        print(f"Loaded with encoding {current_encoding_name}, but data appears empty or invalid. Trying next encoding.")
                        self.spss_load_results["last_exception"] = RuntimeError(f"Encoding {current_encoding_name} resulted in empty or invalid data.")
                except UnicodeDecodeError as ude:
                    self.spss_load_results["last_exception"] = ude
                    print(f"UnicodeDecodeError with encoding {current_encoding_name}: {ude}")
                except pyreadstat.ReadstatError as rse:
                    self.spss_load_results["last_exception"] = rse
                    print(f"PyreadstatError with encoding {current_encoding_name}: {rse}")
                except Exception as e:
                    self.spss_load_results["last_exception"] = e
                    print(f"General error with encoding {current_encoding_name}: {e}")

            self.after(0, self._finish_loading_spss, file_path)

        load_thread = threading.Thread(target=_load_file_in_thread, daemon=True)
        load_thread.start()

    def _finish_loading_spss(self, file_path):
        if hasattr(self, 'loading_dialog') and self.loading_dialog.winfo_exists():
            self.loading_dialog.destroy()

        loaded_successfully = self.spss_load_results["success"]
        used_encoding = self.spss_load_results["used_encoding"]
        last_exception = self.spss_load_results["last_exception"]

        if not loaded_successfully:
            error_message = "Failed to load SPSS file. All attempted encodings failed."
            if last_exception:
                error_message += f"\nLast error ({type(last_exception).__name__}): {str(last_exception)}"
            messagebox.showerror("Error Loading File", error_message)
            self.spss_df_codes, self.spss_df_labels, self.spss_meta, self.spss_variables = None, None, None, []
            self.file_label.configure(text="Failed to load file.")
            return

        self.spss_df_codes = self.spss_load_results["df_codes"]
        self.spss_meta = self.spss_load_results["meta"]

        if self.spss_df_codes is not None and self.spss_meta is not None:
            self.spss_df_labels = self.spss_df_codes.copy()
            if hasattr(self.spss_meta, 'variable_value_labels') and self.spss_meta.variable_value_labels:
                for col_name, label_dict in self.spss_meta.variable_value_labels.items():
                    if col_name in self.spss_df_labels.columns and label_dict:
                        try:
                            mapped_series = self.spss_df_labels[col_name].map(label_dict)
                        except TypeError:
                            try:
                                mapped_series = self.spss_df_labels[col_name].astype(str).map(label_dict)
                            except Exception:
                                mapped_series = pd.Series([None] * len(self.spss_df_labels[col_name]), index=self.spss_df_labels[col_name].index)
                        self.spss_df_labels[col_name] = mapped_series.fillna(self.spss_df_codes[col_name].astype(str)).astype(str)
            else:
                 self.spss_df_labels = self.spss_df_codes.astype(str)
        else:
            self.spss_df_labels = None

        if self.spss_df_codes is not None:
            self.spss_variables = list(self.spss_df_codes.columns)
        else:
            self.spss_variables = []

        self.file_label.configure(text=f"Loaded: {os.path.basename(file_path)} ({len(self.spss_variables)} vars) Encoding: {used_encoding}")

        self._clear_all_selections_and_inputs() # Clear selections as new file is loaded

        messagebox.showinfo("Success", f"SPSS file loaded successfully using encoding '{used_encoding}'.\nCodes and Labels dataframes are prepared.")


    def spss_df_is_loaded(self):
        if self.spss_df_codes is None or self.spss_df_labels is None:
            messagebox.showwarning("No Data", "Please load an SPSS file first.")
            return False
        return True

    def open_demographic_var_selection_dialog(self, field_name_ref, display_widget_ref):
        if not self.spss_df_is_loaded(): return

        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Select SPSS Variable for {field_name_ref}")
        dialog_width = 600; dialog_height = 700
        self._center_window(dialog, dialog_width, dialog_height)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text=f"Choose SPSS variable for '{field_name_ref}':").pack(pady=(10,5))
        search_var = ctk.StringVar()
        search_entry = ctk.CTkEntry(dialog, textvariable=search_var, placeholder_text="Search variable...")
        search_entry.pack(pady=5, padx=10, fill="x")

        list_container_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        list_container_frame.pack(pady=5, padx=10, fill="both", expand=True)

        dialog.variable_list_frame = ctk.CTkScrollableFrame(list_container_frame, height=250)
        # pack() will be handled in _populate_variable_list_actual_demog

        labels_preview_textbox = ctk.CTkTextbox(dialog, height=150, state="disabled", wrap="word")
        labels_preview_textbox.pack(pady=10, padx=10, fill="both", expand=True)

        selected_var_tk = ctk.StringVar()
        current_selection_for_field = self.demographic_var_selections.get(field_name_ref)
        if current_selection_for_field and current_selection_for_field in self.spss_variables:
            selected_var_tk.set(current_selection_for_field)
        else:
            selected_var_tk.set("")

        dialog._search_after_id_demog = None
        dialog._info_label_demog = None

        def _update_labels_preview_demog(selected_spss_var_name):
            labels_preview_textbox.configure(state="normal")
            labels_preview_textbox.delete("1.0", "end")
            if not selected_spss_var_name:
                labels_preview_textbox.insert("1.0", "No variable selected, or variable has no labels, or selection is not currently visible.")
                labels_preview_textbox.configure(state="disabled"); return

            preview_text = f"Value Labels for '{selected_spss_var_name}':\n"
            if self.spss_meta and hasattr(self.spss_meta, 'variable_value_labels') and self.spss_meta.variable_value_labels:
                value_labels = self.spss_meta.variable_value_labels.get(selected_spss_var_name, {})
                if value_labels:
                    for i, (code, label) in enumerate(value_labels.items()):
                        if i >= 15:
                            preview_text += "  ...\n"; break
                        preview_text += f"  {code}: {label}\n"
                else: preview_text += "  (No defined value labels for this variable)"
            else: preview_text += "  (Metadata or value labels not available)"
            labels_preview_textbox.insert("1.0", preview_text)
            labels_preview_textbox.configure(state="disabled")

        def _populate_variable_list_actual_demog(filter_text=""):
            for widget in dialog.variable_list_frame.winfo_children():
                widget.destroy()
            if dialog._info_label_demog and dialog._info_label_demog.winfo_exists():
                dialog._info_label_demog.pack_forget()
                dialog._info_label_demog.destroy()
                dialog._info_label_demog = None

            filtered_vars = []
            if self.spss_variables:
                filter_text_lower = filter_text.lower()
                filtered_vars = [var_name for var_name in self.spss_variables if filter_text_lower in var_name.lower()]

            if not self.spss_variables:
                dialog.variable_list_frame.pack_forget()
                dialog._info_label_demog = ctk.CTkLabel(list_container_frame, text="No SPSS variables loaded.")
                dialog._info_label_demog.pack(anchor="center", pady=10)
                _update_labels_preview_demog(None)
                return
            elif not filtered_vars:
                dialog.variable_list_frame.pack_forget()
                dialog._info_label_demog = ctk.CTkLabel(list_container_frame, text="No variables match your search.")
                dialog._info_label_demog.pack(anchor="center", pady=10)
                _update_labels_preview_demog(None)
                return
            else:
                dialog.variable_list_frame.pack(fill="x", expand=True)


            vars_to_display = filtered_vars[:self.MAX_DISPLAY_ITEMS_IN_DIALOG]
            for var_name in vars_to_display:
                rb = ctk.CTkRadioButton(dialog.variable_list_frame, text=var_name, variable=selected_var_tk, value=var_name,
                                        command=lambda v=var_name: _update_labels_preview_demog(v))
                rb.pack(anchor="w", padx=5, pady=2)

            if len(filtered_vars) > self.MAX_DISPLAY_ITEMS_IN_DIALOG:
                dialog._info_label_demog = ctk.CTkLabel(list_container_frame,
                                     text=f"Showing {self.MAX_DISPLAY_ITEMS_IN_DIALOG} of {len(filtered_vars)} variables. Refine search.",
                                     font=self.dialog_info_font)
                dialog._info_label_demog.pack(anchor="w", padx=5, pady=(5,0), side="bottom")

            current_tk_selection = selected_var_tk.get()
            if current_tk_selection and current_tk_selection in vars_to_display:
                _update_labels_preview_demog(current_tk_selection)
            else:
                _update_labels_preview_demog(None)

        def _on_search_change_demog(*args):
            if dialog._search_after_id_demog is not None:
                dialog.after_cancel(dialog._search_after_id_demog)
            dialog._search_after_id_demog = dialog.after(300, lambda: _populate_variable_list_actual_demog(search_var.get()))

        dialog._current_trace_id_demog = search_var.trace_add("write", _on_search_change_demog)
        _populate_variable_list_actual_demog(search_var.get())

        def on_ok_demog():
            if dialog._search_after_id_demog is not None:
                dialog.after_cancel(dialog._search_after_id_demog)
            chosen_spss_var = selected_var_tk.get()
            if chosen_spss_var and chosen_spss_var in self.spss_variables:
                 self.demographic_var_selections[field_name_ref] = chosen_spss_var
                 display_widget_ref.configure(text=chosen_spss_var)
            else:
                 self.demographic_var_selections[field_name_ref] = None
                 display_widget_ref.configure(text="None selected")
            if hasattr(dialog, '_current_trace_id_demog') and dialog._current_trace_id_demog:
                 search_var.trace_vdelete("w", dialog._current_trace_id_demog)
            dialog.destroy()

        def on_cancel_demog():
            if dialog._search_after_id_demog is not None:
                dialog.after_cancel(dialog._search_after_id_demog)
            if hasattr(dialog, '_current_trace_id_demog') and dialog._current_trace_id_demog:
                 search_var.trace_vdelete("w", dialog._current_trace_id_demog)
            dialog.destroy()

        button_frame = ctk.CTkFrame(dialog); button_frame.pack(pady=(10,10), fill="x")
        ctk.CTkButton(button_frame, text="OK", command=on_ok_demog, width=100).pack(side="left", padx=20, expand=True)
        ctk.CTkButton(button_frame, text="Cancel", command=on_cancel_demog, fg_color="gray", width=100).pack(side="right", padx=20, expand=True)


    def open_multiselect_dialog(self, field_name_ref):
        config = self.fixed_input_fields.get(field_name_ref)
        if not config or config["type"] != "multiselect_dialog_button": return

        options_list = config.get("options", [])
        # Use a temporary set for efficient add/remove and checking existence in the dialog
        temp_selected_values_in_dialog_set = set(config.get("selected_values", []))


        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Select {field_name_ref}")
        dialog_width = 450; dialog_height = 550
        self._center_window(dialog, dialog_width, dialog_height)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text=f"Choose options for '{field_name_ref}':").pack(pady=10)
        search_var_multi = ctk.StringVar()
        search_entry_multi = ctk.CTkEntry(dialog, textvariable=search_var_multi, placeholder_text="Search options...")
        search_entry_multi.pack(pady=5, padx=10, fill="x")

        list_container_frame_multi = ctk.CTkFrame(dialog, fg_color="transparent")
        list_container_frame_multi.pack(pady=5, padx=10, fill="both", expand=True)

        dialog.scroll_frame_multi = ctk.CTkScrollableFrame(list_container_frame_multi, height=350)
        # pack() will be handled in _populate_multiselect_options_list_actual

        dialog._search_after_id_multi = None
        dialog._info_label_multi = None

        def _populate_multiselect_options_list_actual(filter_text=""):
            for widget in dialog.scroll_frame_multi.winfo_children():
                widget.destroy()
            if dialog._info_label_multi and dialog._info_label_multi.winfo_exists():
                dialog._info_label_multi.pack_forget()
                dialog._info_label_multi.destroy()
                dialog._info_label_multi = None

            filtered_options = []
            if options_list:
                filter_text_lower = filter_text.lower()
                filtered_options = [opt for opt in options_list if filter_text_lower in opt.lower()]

            if not options_list:
                dialog.scroll_frame_multi.pack_forget()
                dialog._info_label_multi = ctk.CTkLabel(list_container_frame_multi, text="No options available.")
                dialog._info_label_multi.pack(anchor="center", pady=10)
                return
            elif not filtered_options:
                dialog.scroll_frame_multi.pack_forget()
                dialog._info_label_multi = ctk.CTkLabel(list_container_frame_multi, text="No options match your search.")
                dialog._info_label_multi.pack(anchor="center", pady=10)
                return
            else:
                dialog.scroll_frame_multi.pack(fill="both", expand=True)


            options_to_display = filtered_options[:self.MAX_DISPLAY_ITEMS_IN_DIALOG]

            for option_text in options_to_display:
                var = ctk.IntVar(value=1 if option_text in temp_selected_values_in_dialog_set else 0)

                def create_toggle_command(opt_txt, int_var_ref): # Pass int_var_ref
                    def on_checkbox_toggle():
                        if int_var_ref.get() == 1:
                            temp_selected_values_in_dialog_set.add(opt_txt)
                        else:
                            temp_selected_values_in_dialog_set.discard(opt_txt)
                    return on_checkbox_toggle

                cb = ctk.CTkCheckBox(dialog.scroll_frame_multi, text=option_text, variable=var,
                                     onvalue=1, offvalue=0, command=create_toggle_command(option_text, var))
                cb.pack(anchor="w", padx=10, pady=2)

            if len(filtered_options) > self.MAX_DISPLAY_ITEMS_IN_DIALOG:
                dialog._info_label_multi = ctk.CTkLabel(list_container_frame_multi,
                                     text=f"Showing {self.MAX_DISPLAY_ITEMS_IN_DIALOG} of {len(filtered_options)} options. Refine search.",
                                     font=self.dialog_info_font)
                dialog._info_label_multi.pack(anchor="w", padx=5, pady=(5,0), side="bottom")


        def _on_search_change_multi(*args):
            if dialog._search_after_id_multi is not None:
                dialog.after_cancel(dialog._search_after_id_multi)
            dialog._search_after_id_multi = dialog.after(300, lambda: _populate_multiselect_options_list_actual(search_var_multi.get()))

        dialog._current_trace_id_multi = search_var_multi.trace_add("write", _on_search_change_multi)
        _populate_multiselect_options_list_actual(search_var_multi.get())

        def on_ok_multiselect():
            if dialog._search_after_id_multi is not None:
                dialog.after_cancel(dialog._search_after_id_multi)
            config["selected_values"] = sorted(list(temp_selected_values_in_dialog_set))
            config["var"].set(", ".join(config["selected_values"]))
            if hasattr(dialog, '_current_trace_id_multi') and dialog._current_trace_id_multi:
                search_var_multi.trace_vdelete("w", dialog._current_trace_id_multi)
            dialog.destroy()

        def on_cancel_multiselect():
            if dialog._search_after_id_multi is not None:
                dialog.after_cancel(dialog._search_after_id_multi)
            if hasattr(dialog, '_current_trace_id_multi') and dialog._current_trace_id_multi:
                search_var_multi.trace_vdelete("w", dialog._current_trace_id_multi)
            dialog.destroy()

        button_frame = ctk.CTkFrame(dialog)
        button_frame.pack(pady=10, fill="x")
        ctk.CTkButton(button_frame, text="OK", command=on_ok_multiselect, width=100).pack(side="left", padx=20, expand=True)
        ctk.CTkButton(button_frame, text="Cancel", command=on_cancel_multiselect, fg_color="gray", width=100).pack(side="right", padx=20, expand=True)


    def get_variable_display_info(self, var_name):
        min_scale, max_scale = None, None
        has_valid_scale = False
        if self.spss_meta and hasattr(self.spss_meta, 'variable_value_labels') and self.spss_meta.variable_value_labels:
            value_labels_for_var = self.spss_meta.variable_value_labels.get(var_name)
            if value_labels_for_var and isinstance(value_labels_for_var, dict) and value_labels_for_var:
                try:
                    codes = sorted([k for k in value_labels_for_var.keys() if isinstance(k, (int, float))])
                    if codes:
                        min_scale, max_scale = codes[0], codes[-1]
                        has_valid_scale = True
                except TypeError:
                    pass # Handles non-numeric codes if any, though pyreadstat usually gives numeric
        return var_name, min_scale, max_scale, has_valid_scale


    def open_channel_var_selection_dialog(self, channel_index):
        if not self.spss_df_is_loaded(): return

        dialog = ctk.CTkToplevel(self)
        dialog_title = f"Select Variables for: {self.channel_custom_names[channel_index]}"
        dialog.title(dialog_title)
        dialog_width = 350 # Adjusted width slightly
        dialog_height = 750
        self._center_window(dialog, dialog_width, dialog_height)
        dialog.grab_set()

        top_dialog_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        top_dialog_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(top_dialog_frame, text=f"Select for: {self.channel_custom_names[channel_index]}").pack(side="left")

        _channel_selected_count_var = ctk.StringVar(value="Selected: 0")
        ctk.CTkLabel(top_dialog_frame, textvariable=_channel_selected_count_var).pack(side="right", padx=10)

        search_var_channel = ctk.StringVar()
        search_entry_channel = ctk.CTkEntry(dialog, textvariable=search_var_channel, placeholder_text="Search variable...")
        search_entry_channel.pack(pady=5, padx=10, fill="x")

        list_container_frame_channel = ctk.CTkFrame(dialog, fg_color="transparent")
        list_container_frame_channel.pack(pady=5, padx=10, fill="both", expand=True)

        dialog.variable_list_frame_channel = ctk.CTkScrollableFrame(list_container_frame_channel, height=380)
        # pack() handled in _filter_and_display_channel_vars_actual

        channel_labels_preview_textbox = ctk.CTkTextbox(dialog, height=150, state="disabled", wrap="word")
        channel_labels_preview_textbox.pack(pady=(5,10), padx=10, fill="x")

        # Temporary state for this dialog instance, initialized from main app state
        # Stores {'var_name': {'selected': BooleanVar, 'reverse': BooleanVar, 'has_scale': bool, ...}}
        dialog.temp_channel_vars_state = {}
        initial_selection_for_channel = self.channel_selections.get(channel_index, [])
        for var_info in initial_selection_for_channel:
            var_name = var_info['name']
            _, min_s, max_s, has_valid_s = self.get_variable_display_info(var_name)
            dialog.temp_channel_vars_state[var_name] = {
                'selected_tk': ctk.BooleanVar(value=True), # If it's in initial_selection, it's selected
                'reverse_tk': ctk.BooleanVar(value=var_info.get('reverse', False)),
                'min_scale': min_s, 'max_scale': max_s,
                'has_scale': has_valid_s
            }

        dialog._search_after_id_channel = None
        dialog._info_label_channel = None
        dialog._last_previewed_var_channel = None # Store the var name whose labels are shown


        def _update_selected_count_channel():
            count = 0
            # Iterate through all SPSS variables, check their state in temp_channel_vars_state
            for var_name in self.spss_variables:
                state = dialog.temp_channel_vars_state.get(var_name)
                if state and state['selected_tk'].get():
                    count +=1
            _channel_selected_count_var.set(f"Selected: {count}")


        def _update_channel_labels_preview(selected_spss_var_name):
            dialog._last_previewed_var_channel = selected_spss_var_name # Update last previewed
            channel_labels_preview_textbox.configure(state="normal")
            channel_labels_preview_textbox.delete("1.0", "end")
            if not selected_spss_var_name:
                channel_labels_preview_textbox.insert("1.0", "Hover over or select a variable name to see its value labels.")
                channel_labels_preview_textbox.configure(state="disabled"); return

            preview_text = f"Value Labels for '{selected_spss_var_name}':\n"
            if self.spss_meta and hasattr(self.spss_meta, 'variable_value_labels') and self.spss_meta.variable_value_labels:
                value_labels = self.spss_meta.variable_value_labels.get(selected_spss_var_name, {})
                if value_labels:
                    for i, (code, label) in enumerate(value_labels.items()):
                        if i >= 15: preview_text += "  ...\n"; break
                        preview_text += f"  {code}: {label}\n"
                else: preview_text += "  (No defined value labels for this variable)"
            else: preview_text += "  (Metadata or value labels not available)"
            channel_labels_preview_textbox.insert("1.0", preview_text)
            channel_labels_preview_textbox.configure(state="disabled")

        def _on_channel_var_toggled_factory(var_name):
            def _actual_toggle():
                # This function is called when a checkbox is clicked.
                # The BooleanVar is already updated by CTk.
                _update_selected_count_channel()
                # Optionally, update preview if selection changes focus,
                # but hover might be better for preview.
                # For now, rely on hover/focus for preview update
            return _actual_toggle

        def _on_item_enter_factory(var_name): # For hover preview
            def _actual_enter(event):
                _update_channel_labels_preview(var_name)
            return _actual_enter

        def _on_item_leave_factory(): # For hover preview
            def _actual_leave(event):
                 # Optionally, revert to a "default" preview or last selected if needed
                 # For now, if mouse leaves, it keeps showing the last hovered var's labels
                 # or clear it if preferred:
                 # _update_channel_labels_preview(dialog._last_previewed_var_channel or None)
                 pass
            return _actual_leave


        def _filter_and_display_channel_vars_actual(filter_text=""):
            for widget in dialog.variable_list_frame_channel.winfo_children():
                widget.destroy()
            if dialog._info_label_channel and dialog._info_label_channel.winfo_exists():
                dialog._info_label_channel.pack_forget()
                dialog._info_label_channel.destroy()
                dialog._info_label_channel = None

            filtered_vars = []
            if self.spss_variables:
                filter_text_lower = filter_text.lower()
                filtered_vars = [var for var in self.spss_variables if filter_text_lower in var.lower()]

            if not self.spss_variables:
                dialog.variable_list_frame_channel.pack_forget()
                dialog._info_label_channel = ctk.CTkLabel(list_container_frame_channel, text="No SPSS variables loaded.")
                dialog._info_label_channel.pack(anchor="center", pady=10)
                _update_channel_labels_preview(None)
                _update_selected_count_channel()
                return
            elif not filtered_vars:
                dialog.variable_list_frame_channel.pack_forget()
                dialog._info_label_channel = ctk.CTkLabel(list_container_frame_channel, text="No variables match your search.")
                dialog._info_label_channel.pack(anchor="center", pady=10)
                _update_channel_labels_preview(None)
                _update_selected_count_channel() # Count should be 0 if nothing matches
                return
            else:
                dialog.variable_list_frame_channel.pack(fill="both", expand=True)

            vars_to_display = filtered_vars[:self.MAX_DISPLAY_ITEMS_IN_DIALOG]

            for var_name in vars_to_display:
                # Ensure state exists for this var_name, create if not (e.g., if it wasn't initially selected)
                if var_name not in dialog.temp_channel_vars_state:
                    _, min_s, max_s, has_valid_s = self.get_variable_display_info(var_name)
                    dialog.temp_channel_vars_state[var_name] = {
                        'selected_tk': ctk.BooleanVar(value=False),
                        'reverse_tk': ctk.BooleanVar(value=False),
                        'min_scale': min_s, 'max_scale': max_s,
                        'has_scale': has_valid_s
                    }
                current_var_state = dialog.temp_channel_vars_state[var_name]

                var_item_frame = ctk.CTkFrame(dialog.variable_list_frame_channel) # Create frame for each item
                var_item_frame.pack(fill="x", pady=1, padx=1) # Small padding

                cb_select = ctk.CTkCheckBox(var_item_frame, text=var_name,
                                            variable=current_var_state['selected_tk'],
                                            onvalue=True, offvalue=False,
                                            command=_on_channel_var_toggled_factory(var_name),
                                            width=10) # Adjust width as needed, expand will take over
                cb_select.pack(side="left", padx=(0,5), expand=True, fill="x")

                # Bind hover events for preview
                var_item_frame.bind("<Enter>", _on_item_enter_factory(var_name))
                # var_item_frame.bind("<Leave>", _on_item_leave_factory()) # Can be noisy
                cb_select.bind("<Enter>", _on_item_enter_factory(var_name)) # Also on checkbox itself
                # cb_select.bind("<Leave>", _on_item_leave_factory())

                cb_reverse = ctk.CTkCheckBox(var_item_frame, text="กลับสเกล", width=80, # Reduced width
                                             variable=current_var_state['reverse_tk'],
                                             onvalue=True, offvalue=False)
                cb_reverse.pack(side="left", padx=2)
                if not current_var_state['has_scale']:
                    cb_reverse.configure(state="disabled")


            if len(filtered_vars) > self.MAX_DISPLAY_ITEMS_IN_DIALOG:
                dialog._info_label_channel = ctk.CTkLabel(list_container_frame_channel,
                                     text=f"Showing {self.MAX_DISPLAY_ITEMS_IN_DIALOG} of {len(filtered_vars)}. Refine search.",
                                     font=self.dialog_info_font)
                dialog._info_label_channel.pack(anchor="w", padx=5, pady=(5,0), side="bottom")

            _update_selected_count_channel() # Update count after displaying
            # Determine what to preview initially or after filter
            # If last previewed var is still in the list, keep it. Else, clear or pick first selected.
            if dialog._last_previewed_var_channel and dialog._last_previewed_var_channel in vars_to_display:
                _update_channel_labels_preview(dialog._last_previewed_var_channel)
            else:
                # Find first selected and visible var to preview
                first_selected_visible = None
                for v_name in vars_to_display:
                    state = dialog.temp_channel_vars_state.get(v_name)
                    if state and state['selected_tk'].get():
                        first_selected_visible = v_name
                        break
                _update_channel_labels_preview(first_selected_visible)


        def _on_search_change_channel(*args):
            if dialog._search_after_id_channel is not None:
                dialog.after_cancel(dialog._search_after_id_channel)
            dialog._search_after_id_channel = dialog.after(300, lambda: _filter_and_display_channel_vars_actual(search_var_channel.get()))

        dialog._current_trace_id_channel = search_var_channel.trace_add("write", _on_search_change_channel)
        _filter_and_display_channel_vars_actual(search_var_channel.get()) # Initial population

        def on_ok_channel():
            if dialog._search_after_id_channel is not None:
                dialog.after_cancel(dialog._search_after_id_channel)

            new_selection_for_this_channel = []
            # Iterate all SPSS variables to capture selections even if not currently displayed due to filtering
            for var_name in self.spss_variables:
                state_info = dialog.temp_channel_vars_state.get(var_name)
                if state_info and state_info['selected_tk'].get():
                    new_selection_for_this_channel.append({
                        'name': var_name,
                        'reverse': state_info['reverse_tk'].get() and state_info['has_scale'],
                        'min_scale': state_info['min_scale'],
                        'max_scale': state_info['max_scale'],
                        'display_text_short': var_name # Or a shorter version if available
                    })
            self.channel_selections[channel_index] = new_selection_for_this_channel
            self.update_channel_display_text(channel_index)
            if hasattr(dialog, '_current_trace_id_channel') and dialog._current_trace_id_channel:
                search_var_channel.trace_vdelete("w", dialog._current_trace_id_channel)
            dialog.destroy()

        def on_cancel_channel():
            if dialog._search_after_id_channel is not None:
                dialog.after_cancel(dialog._search_after_id_channel)
            if hasattr(dialog, '_current_trace_id_channel') and dialog._current_trace_id_channel:
                search_var_channel.trace_vdelete("w", dialog._current_trace_id_channel)
            dialog.destroy()

        button_frame = ctk.CTkFrame(dialog); button_frame.pack(pady=10, fill="x")
        ctk.CTkButton(button_frame, text="OK", command=on_ok_channel, width=100).pack(side="left", padx=20, expand=True)
        ctk.CTkButton(button_frame, text="Cancel", command=on_cancel_channel, fg_color="gray", width=100).pack(side="right", padx=20, expand=True)


    def update_channel_display_text(self, channel_index):
        display_widget = self.channel_widgets[channel_index]["display"]
        selected_vars_info_list = self.channel_selections.get(channel_index, [])
        texts_to_display = []
        for var_info in selected_vars_info_list:
            display_name = var_info['name']
            if var_info.get('reverse', False):
                display_name += " (R)"
            texts_to_display.append(display_name)

        display_widget.configure(state="normal")
        display_widget.delete("1.0", "end")
        if texts_to_display:
            display_widget.insert("1.0", ", ".join(texts_to_display))
        else:
            display_widget.insert("1.0", "None selected")
        display_widget.configure(state="disabled")


    def generate_index_sheet_data(self):
        index_data_row = {}
        for field_name, config in self.fixed_input_fields.items():
            value = ""
            if config["type"] == "multiselect_dialog_button":
                value = ", ".join(config.get("selected_values",[]))
            elif "var" in config and config["var"] is not None:
                value = config["var"].get()

            if field_name == "Timestamp" and value == "Auto (on export)":
                index_data_row[field_name] = "Auto"
            else:
                index_data_row[field_name] = value if value and value != "None" else ""

        for field_name in self.demographic_fields:
            spss_var_name = self.demographic_var_selections.get(field_name)
            index_data_row[field_name] = spss_var_name if spss_var_name else ""

        for ch_idx in range(self.num_channels):
            channel_header_name = self.channel_custom_names[ch_idx]
            vars_info = self.channel_selections.get(ch_idx, [])
            texts = [var_info['name'] for var_info in vars_info]
            index_data_row[channel_header_name] = ", ".join(texts) if texts else ""
        return pd.DataFrame([index_data_row])


    def _prepare_export_dataframes(self):
        if not self.spss_df_is_loaded(): return None, None, None

        output_data_channels_dict, max_len_from_channels = {}, 0
        if self.spss_df_codes is not None:
            for ch_idx in range(self.num_channels):
                channel_name_export = self.channel_custom_names[ch_idx]
                vars_info_for_channel = self.channel_selections.get(ch_idx, [])
                current_channel_data_list = []

                for var_info in vars_info_for_channel:
                    var_name_export = var_info['name']
                    if var_name_export not in self.spss_df_codes.columns: continue
                    series_data = self.spss_df_codes[var_name_export].copy()
                    if var_info.get('reverse') and var_info['min_scale'] is not None and var_info['max_scale'] is not None:
                        try: # Add try-except for robust conversion
                            min_s, max_s = float(var_info['min_scale']), float(var_info['max_scale'])
                            series_data = series_data.apply(
                                lambda val: (min_s + max_s) - float(val)
                                if pd.notna(val) and isinstance(val, (int, float)) # Check type before conversion
                                else val
                            )
                        except (ValueError, TypeError) as e:
                            print(f"Warning: Could not reverse scale for {var_name_export} due to non-numeric scale values or data: {e}")

                    current_channel_data_list.extend(series_data.tolist())
                output_data_channels_dict[channel_name_export] = pd.Series(current_channel_data_list)
                max_len_from_channels = max(max_len_from_channels, len(current_channel_data_list))

        if max_len_from_channels == 0 and self.spss_df_codes is not None and not self.spss_df_codes.empty:
             max_len_from_channels = len(self.spss_df_codes)
        elif max_len_from_channels == 0: # If still 0, means no channels selected and SPSS df might be empty
            max_len_from_channels = 1 # Default to 1 row for fixed inputs if everything else is empty


        output_data_fixed_dict, timestamp_value_for_rawdata = {}, ""
        for field_name, config in self.fixed_input_fields.items():
            value_to_use = ""
            if config["type"] == "multiselect_dialog_button":
                value_to_use = ", ".join(config.get("selected_values",[]))
            elif "var" in config and config["var"] is not None:
                value_to_use = config["var"].get()

            if field_name == "Timestamp" and value_to_use == "Auto (on export)":
                timestamp_value_for_rawdata = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                value_to_use = timestamp_value_for_rawdata
            elif field_name == "Timestamp": # if user provided a timestamp
                timestamp_value_for_rawdata = value_to_use if value_to_use is not None else datetime.now().strftime("%Y-%m-%d %H:%M:%S") # Fallback if somehow empty
            output_data_fixed_dict[field_name] = pd.Series([value_to_use if value_to_use is not None else ""] * max_len_from_channels, dtype=str)

        output_data_demographics_dict = {}
        if self.spss_df_labels is not None and not self.spss_df_labels.empty:
            num_cases_spss = len(self.spss_df_labels)
            for field_name in self.demographic_fields:
                spss_var_name = self.demographic_var_selections.get(field_name)
                if spss_var_name and spss_var_name in self.spss_df_labels.columns:
                    demographic_series = self.spss_df_labels[spss_var_name].astype(str).fillna('')
                    if num_cases_spss > 0 and max_len_from_channels > 0 :
                        if max_len_from_channels >= num_cases_spss: # Replicate if rawdata is longer
                            num_repeats = (max_len_from_channels + num_cases_spss - 1) // num_cases_spss
                            replicated_data = pd.concat([demographic_series] * num_repeats, ignore_index=True)[:max_len_from_channels]
                        else: # Truncate if rawdata is shorter
                            replicated_data = demographic_series[:max_len_from_channels]
                    elif max_len_from_channels > 0: # No SPSS data but rawdata rows exist
                         replicated_data = pd.Series([''] * max_len_from_channels, dtype=str)
                    else: # No rawdata rows
                        replicated_data = pd.Series([], dtype=str)
                    output_data_demographics_dict[field_name] = replicated_data
                else:
                    output_data_demographics_dict[field_name] = pd.Series([''] * max_len_from_channels, dtype=str)
        else: # No SPSS labels dataframe
            for field_name in self.demographic_fields:
                output_data_demographics_dict[field_name] = pd.Series([''] * max_len_from_channels, dtype=str)

        final_df_list_rawdata = []
        expected_headers_ordered = list(self.fixed_input_fields.keys()) + self.demographic_fields + self.channel_custom_names

        # Add fixed inputs ensuring all columns exist
        df_fixed = pd.DataFrame(output_data_fixed_dict)
        for col in self.fixed_input_fields.keys():
            if col not in df_fixed.columns: df_fixed[col] = pd.Series([''] * max_len_from_channels, dtype=str)
        final_df_list_rawdata.append(df_fixed[[col for col in self.fixed_input_fields.keys()]])


        # Add demographics ensuring all columns exist
        df_demographics = pd.DataFrame(output_data_demographics_dict)
        for col in self.demographic_fields:
            if col not in df_demographics.columns: df_demographics[col] = pd.Series([''] * max_len_from_channels, dtype=str)
        final_df_list_rawdata.append(df_demographics[self.demographic_fields])


        # Add channels ensuring all columns exist
        df_channels = pd.DataFrame(output_data_channels_dict)
        for col in self.channel_custom_names:
            if col not in df_channels.columns: df_channels[col] = pd.Series([''] * max_len_from_channels, dtype=str)
        final_df_list_rawdata.append(df_channels[self.channel_custom_names])


        df_rawdata = pd.DataFrame(columns=expected_headers_ordered).astype(str) # Start with empty df with correct order
        if final_df_list_rawdata:
            # Concatenate ensuring correct column order and handling of empty parts
            temp_rawdata = pd.concat(final_df_list_rawdata, axis=1)
            # Reindex to ensure all expected columns are present and in order
            df_rawdata = temp_rawdata.reindex(columns=expected_headers_ordered, fill_value='').astype(str)

        df_rawdata = df_rawdata.fillna('').astype(str)

        df_index = self.generate_index_sheet_data()
        if "Timestamp" in df_index.columns and df_index.loc[0, "Timestamp"] == "Auto" and timestamp_value_for_rawdata:
             df_index.loc[0, "Timestamp"] = timestamp_value_for_rawdata
        df_index = df_index.fillna('').astype(str)
        return df_index, df_rawdata, timestamp_value_for_rawdata


    def export_to_excel(self):
        df_index, df_rawdata, _ = self._prepare_export_dataframes()
        if df_index is None or df_rawdata is None:
            messagebox.showwarning("Export Error", "Could not prepare data for export. Was an SPSS file loaded?")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel files", "*.xlsx"),("All files", "*.*")), title="Save Exported Data As (Excel)")
        if not save_path: return
        if os.path.splitext(save_path)[1].lower() != ".xlsx":
            save_path = os.path.splitext(save_path)[0] + ".xlsx"

        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df_index.to_excel(writer, sheet_name='Sheet Index', index=False)
                df_rawdata.to_excel(writer, sheet_name='RawdataNorm', index=False)
            messagebox.showinfo("Success", f"Data exported to Excel successfully!\n{save_path}")
        except Exception as e:
            messagebox.showerror("Excel Export Error", str(e))


    def fetch_total_jobs_from_gsheet_threaded(self):
        self.total_jobs_var.set("Total Jobs: Loading...")
        thread = threading.Thread(target=self.fetch_total_jobs_from_gsheet, daemon=True)
        thread.start()

    def fetch_total_jobs_from_gsheet(self):
        try:
            if not hasattr(self, 'SERVICE_ACCOUNT_FILE_PATH') or not self.SERVICE_ACCOUNT_FILE_PATH:
                self.total_jobs_var.set("Total Jobs: Key Path N/A")
                print("Error: SERVICE_ACCOUNT_FILE_PATH is not defined in the class.")
                return
            if not hasattr(self, 'INDEX_SHEET_NAME_FOR_COUNT') or not self.INDEX_SHEET_NAME_FOR_COUNT:
                self.total_jobs_var.set("Total Jobs: Index Sheet N/A")
                print("Error: INDEX_SHEET_NAME_FOR_COUNT is not defined in the class.")
                return
            if not hasattr(self, 'TARGET_GOOGLE_SHEET_URL') or not self.TARGET_GOOGLE_SHEET_URL:
                self.total_jobs_var.set("Total Jobs: GSheet URL N/A")
                print("Error: TARGET_GOOGLE_SHEET_URL is not defined in the class.")
                return
            if not hasattr(self, 'UNIQUE_JOB_IDENTIFIER_COLUMN') or not self.UNIQUE_JOB_IDENTIFIER_COLUMN:
                self.total_jobs_var.set("Total Jobs: Unique Col N/A")
                print("Error: UNIQUE_JOB_IDENTIFIER_COLUMN is not defined in the class.")
                return

            if not os.path.exists(self.SERVICE_ACCOUNT_FILE_PATH):
                self.total_jobs_var.set("Total Jobs: Key File Err")
                print(f"Service account key file not found: {self.SERVICE_ACCOUNT_FILE_PATH}")
                return

            scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
            creds = Credentials.from_service_account_file(self.SERVICE_ACCOUNT_FILE_PATH, scopes=scopes)
            client = gspread.authorize(creds)
            spreadsheet = client.open_by_url(self.TARGET_GOOGLE_SHEET_URL)
            worksheet = spreadsheet.worksheet(self.INDEX_SHEET_NAME_FOR_COUNT)
            all_records = worksheet.get_all_records() # Fetches data as list of dicts

            if not all_records:
                self.total_jobs_var.set("Total Jobs: 0")
                return

            unique_job_ids = set()
            for record in all_records:
                job_id = record.get(self.UNIQUE_JOB_IDENTIFIER_COLUMN)
                # Ensure job_id is treated as string and stripped, handle None
                if job_id is not None and str(job_id).strip() != "":
                    unique_job_ids.add(str(job_id).strip())
            num_unique_jobs = len(unique_job_ids)
            self.total_jobs_var.set(f"Total Jobs: {num_unique_jobs}")

        except gspread.exceptions.APIError as api_e:
            if "Rate Limit Exceeded" in str(api_e) or "RESOURCE_EXHAUSTED" in str(api_e).upper():
                self.total_jobs_var.set("Total Jobs: Rate Limit")
                print(f"Google Sheets API Rate Limit Exceeded during fetch_total_jobs: {str(api_e)}")
            else:
                self.total_jobs_var.set("Total Jobs: API Error")
                print(f"Google Sheets API Error during fetch_total_jobs: {str(api_e)}")
        except gspread.exceptions.SpreadsheetNotFound:
            self.total_jobs_var.set("Total Jobs: Sheet NF")
            print(f"Google Sheet not found or permission denied. URL: {self.TARGET_GOOGLE_SHEET_URL}")
        except gspread.exceptions.WorksheetNotFound:
            self.total_jobs_var.set(f"Total Jobs: Tab NF") # Tab (Worksheet) Not Found
            print(f"Worksheet '{self.INDEX_SHEET_NAME_FOR_COUNT}' not found in the Google Sheet.")
        except FileNotFoundError: # For self.SERVICE_ACCOUNT_FILE_PATH
             self.total_jobs_var.set("Total Jobs: Key File NF")
             print(f"Service account key file not found: {self.SERVICE_ACCOUNT_FILE_PATH}")
        except Exception as e:
            self.total_jobs_var.set("Total Jobs: Error")
            print(f"Error fetching total jobs from Google Sheet: {type(e).__name__} - {str(e)}")


    def save_to_google_sheet(self):
        df_index, df_rawdata, timestamp_val = self._prepare_export_dataframes() # Get timestamp if generated
        if df_index is None or df_rawdata is None:
            messagebox.showwarning("Save Error", "Could not prepare data for Google Sheets. Was an SPSS file loaded?")
            return

        # Check if essential identifying information is missing or default
        project_num = self.fixed_input_fields["Project #"]["var"].get().strip()
        project_name = self.fixed_input_fields["Project Name"]["var"].get().strip()
        username = self.fixed_input_fields["Username"]["var"].get()

        if not project_num or not project_name or username == self.username_placeholder:
            warning_msg = "Please provide at least:\n"
            if not project_num: warning_msg += "- Project #\n"
            if not project_name: warning_msg += "- Project Name\n"
            if username == self.username_placeholder: warning_msg += "- Username\n"
            messagebox.showwarning("Missing Information", warning_msg.strip())
            return

        # Further check if any meaningful data is actually selected beyond defaults
        fixed_inputs_are_meaningful = False
        for field_name, config in self.fixed_input_fields.items():
            if field_name in ["Project #", "Project Name", "Username"]: continue # Already checked
            if config.get("hidden"): continue

            current_value = config["var"].get()
            default_value = config.get("default", "")
            if config["type"] == "multiselect_dialog_button":
                if config.get("selected_values"): # If list is not empty
                    fixed_inputs_are_meaningful = True; break
            elif config["type"] == "combobox" and not default_value and config.get("options"):
                default_value = config["options"][0] # Default for combobox is first option if 'default' key is empty
            
            if current_value != default_value and current_value != "":
                fixed_inputs_are_meaningful = True; break
        
        demographics_selected = any(val for val in self.demographic_var_selections.values())
        channels_selected = any(sel_list for sel_list in self.channel_selections.values())

        if not fixed_inputs_are_meaningful and not demographics_selected and not channels_selected:
            if not messagebox.askyesno("Confirm Save", "It seems only Project #, Name, and Username are filled. "
                                    "No other project info, demographics, or channel variables are selected. "
                                    "Do you still want to save this to Google Sheets?"):
                return

        sheet_identifier = self.TARGET_GOOGLE_SHEET_URL
        try:
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            if not os.path.exists(self.SERVICE_ACCOUNT_FILE_PATH):
                messagebox.showerror("Config Error", f"Service account key file not found:\n{self.SERVICE_ACCOUNT_FILE_PATH}\nPlease place it in the correct path or update the path in the script."); return

            creds = Credentials.from_service_account_file(self.SERVICE_ACCOUNT_FILE_PATH, scopes=scopes)
            client = gspread.authorize(creds)
            spreadsheet = None
            try:
                if "docs.google.com/spreadsheets" in sheet_identifier.lower():
                    spreadsheet = client.open_by_url(sheet_identifier)
                else: # Assume it's a sheet ID (key)
                    spreadsheet = client.open_by_key(sheet_identifier) # Requires only sheet ID
            except gspread.exceptions.SpreadsheetNotFound:
                messagebox.showerror("Error", f"Google Sheet not found or permission denied.\nID/URL: {sheet_identifier}\nEnsure it's shared with the service account email:\n{creds.service_account_email}"); return
            except Exception as e_gs_open: # Catch other gspread or network errors
                messagebox.showerror("Error", f"Could not open Google Sheet: {type(e_gs_open).__name__} - {str(e_gs_open)}"); return

            def prepare_df_for_gspread(input_df):
                if input_df is None: return pd.DataFrame() # Return empty DataFrame if None
                df_copy = input_df.copy()
                # Convert all to object first to handle mixed types before fillna and astype(str)
                df_processed = df_copy.astype(object).fillna('')
                for col in df_processed.columns:
                    # Explicitly replace pandas/numpy specific NA representations if any survive
                    df_processed[col] = df_processed[col].astype(str).replace(['nan', 'NaN', 'None', 'none', '<NA>'], '', regex=False)
                return df_processed.fillna('') # Final safety net

            df_index_clean = prepare_df_for_gspread(df_index)
            df_rawdata_clean = prepare_df_for_gspread(df_rawdata)

            def append_data_to_sheet(worksheet_title, dataframe_to_append):
                try:
                    worksheet = spreadsheet.worksheet(worksheet_title)
                    data_values_to_append = dataframe_to_append.values.tolist() # Headerless data
                    headers_to_append = dataframe_to_append.columns.tolist()   # Column names

                    # Fetch current headers only if we have headers to append and data to append
                    current_headers = []
                    if headers_to_append and not dataframe_to_append.empty:
                        try:
                            current_headers = worksheet.row_values(1)
                        except gspread.exceptions.APIError as e: # Handle empty sheet case
                            if "exceeds grid limits" in str(e).lower() or "empty" in str(e).lower():
                                current_headers = [] # Sheet is likely empty or has no row 1
                            else: raise # Re-raise other API errors
                        except Exception: # Catch any other error during header fetch
                            current_headers = []

                    # Append headers if:
                    # 1. Sheet is empty (no current_headers) AND we have headers_to_append.
                    # 2. Sheet has headers, but they don't match headers_to_append (should ideally not happen if structure is fixed).
                    if (not current_headers and headers_to_append) or \
                    (current_headers and headers_to_append and current_headers != headers_to_append and not dataframe_to_append.empty):
                        if not current_headers: # Only append if sheet was truly empty
                            worksheet.append_row(headers_to_append, value_input_option='USER_ENTERED')
                            print(f"Appended headers to '{worksheet_title}'.")
                        # If headers differ and sheet wasn't empty, this is an issue.
                        # For now, we assume if headers exist, they are correct.
                        # A more robust solution would involve aligning columns or erroring.

                    # Append data rows if any exist
                    if data_values_to_append:
                        # Avoid appending only a header row if it's already there
                        if worksheet_title == self.INDEX_SHEET_NAME_FOR_COUNT and \
                        len(data_values_to_append) == 1 and \
                        headers_to_append == data_values_to_append[0] and \
                        current_headers == headers_to_append:
                            print(f"Data for '{worksheet_title}' is identical to existing headers. Skipping row append.")
                        else:
                            worksheet.append_rows(data_values_to_append, value_input_option='USER_ENTERED')
                            print(f"Appended {len(data_values_to_append)} data rows to '{worksheet_title}'.")
                    else:
                        print(f"No data rows to append to '{worksheet_title}'.")

                except gspread.exceptions.WorksheetNotFound:
                    messagebox.showerror("Error", f"Sheet '{worksheet_title}' not found in the Google Spreadsheet. Please create it or check the name.")
                except Exception as e_sheet:
                    messagebox.showerror("Error Writing to Sheet", f"Error writing to sheet '{worksheet_title}':\n{type(e_sheet).__name__}: {str(e_sheet)}")

            # Append to Index sheet first
            if not df_index_clean.empty:
                append_data_to_sheet(self.INDEX_SHEET_NAME_FOR_COUNT, df_index_clean)
            else:
                print(f"Index DataFrame ('{self.INDEX_SHEET_NAME_FOR_COUNT}') is empty. Skipping append for it.")

            # Append to RawdataNorm sheet
            if not df_rawdata_clean.empty:
                append_data_to_sheet("RawdataNorm", df_rawdata_clean)
            else:
                print("Rawdata DataFrame ('RawdataNorm') is empty. Skipping append for it.")

            messagebox.showinfo("Success", "Data successfully processed for Google Sheets!\nCheck console for append details.")
            
            # Clear only channel selections after successful save (preserve Project Info & Demographics)
            self._clear_all_selections_and_inputs(clear_project_info=False, clear_demographics=False)
            self.fetch_total_jobs_from_gsheet_threaded() # Refresh job count

        except FileNotFoundError: # For SERVICE_ACCOUNT_FILE_PATH
            messagebox.showerror("Config Error", f"Service Account Key File\n'{self.SERVICE_ACCOUNT_FILE_PATH}'\nnot found. Please configure the path in the script.")
        except Exception as e_auth: # Catch-all for other auth/gspread setup errors
            messagebox.showerror("Google Sheets Error", f"An error occurred with Google Sheets setup:\n{type(e_auth).__name__}: {str(e_auth)}")
def run_this_app(working_dir=None):
    print(f"--- SPSS_EXPORTER_INFO: Starting 'SpssExporterApp' via run_this_app() ---")
    # Optional: Change working directory if needed (e.g., for finding Test3.json)
    # if working_dir and os.path.exists(working_dir):
    #     os.chdir(working_dir)
    #     print(f"--- SPSS_EXPORTER_INFO: Changed working directory to: {os.getcwd()} ---")
    # elif working_dir:
    #     print(f"--- SPSS_EXPORTER_WARNING: Provided working_dir '{working_dir}' does not exist. ---")

    try:
        app = SpssExporterApp()
        app.mainloop()
    except Exception as e:
        print(f"SPSS_EXPORTER_ERROR: An error occurred during SpssExporterApp execution: {e}")
        import traceback
        traceback.print_exc() # Print full traceback for debugging
        root_temp = None
        try:
            # Try to show error in a messagebox if GUI can still be partially created
            if tk._default_root and tk._default_root.winfo_exists():
                parent_window = tk._default_root
            else:
                root_temp = tk.Tk()
                root_temp.withdraw() # Hide the empty root window
                parent_window = root_temp
            messagebox.showerror("Application Error (SPSS Exporter)",
                                 f"An unexpected error occurred and the application has to close:\n{e}\n\nCheck the console for more details.",
                                 parent=parent_window)
        except Exception as e_msgbox:
             print(f"SPSS_EXPORTER_ERROR: Could not display error in messagebox: {e_msgbox}")
        finally:
            if root_temp:
                root_temp.destroy()
        sys.exit(f"Error running SpssExporterApp: {e}")

if __name__ == "__main__":
    # To make it easier to run if Test3.json is in the same directory as the script:
    # script_dir = os.path.dirname(os.path.abspath(__file__))
    print("--- Running SpssExporterApp.py directly for testing ---")
    # print(f"--- Script directory: {script_dir} ---")
    # print(f"--- Current working directory before run: {os.getcwd()} ---")
    # run_this_app(working_dir=script_dir) # Pass script_dir if you want to ensure it runs from there
    run_this_app()
    print("--- Finished direct execution of SpssExporterApp.py ---")