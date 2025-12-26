import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Listbox, Scrollbar
import pandas as pd
import numpy as np
import io
import copy

# FIX: Explicitly set Matplotlib backend for Tkinter
import matplotlib
matplotlib.use('TkAgg')

import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class PSMApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Price Sensitivity Meter (PSM) Analysis Tool v2.7")
        self.root.geometry("1400x850")

        self.df = None
        self.filtered_df = None
        self.psm_data = None
        self.current_fig = None
        self.current_intersections = {} 
        self.filter_widgets = []
        self.export_queue = {} 

        # --- Main Frames ---
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(side=tk.TOP, fill=tk.X)
        
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.chart_frame = ttk.Frame(main_frame)
        self.chart_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        export_panel_frame = ttk.Frame(main_frame, padding="10")
        export_panel_frame.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Top Control Widgets ---
        # NEW: Added "Get Template" button
        self.btn_get_template = ttk.Button(top_frame, text="Get Template", command=self.get_template)
        self.btn_get_template.pack(side=tk.LEFT, padx=(0, 10), pady=5)
        
        self.btn_load = ttk.Button(top_frame, text="Load Excel File", command=self.load_excel)
        self.btn_load.pack(side=tk.LEFT, padx=5, pady=5)

        self.file_label = ttk.Label(top_frame, text="No file loaded")
        self.file_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.filter_frame = ttk.Frame(top_frame, padding="5")
        self.filter_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        
        # --- Chart Area ---
        self.fig, self.ax = plt.subplots(figsize=(10, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # --- Export Panel Widgets ---
        ttk.Label(export_panel_frame, text="Export Queue", font=("Segoe UI", 10, "bold")).pack(pady=5)
        
        listbox_frame = ttk.Frame(export_panel_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True)
        self.queue_listbox = Listbox(listbox_frame, height=15)
        self.queue_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = Scrollbar(listbox_frame, orient="vertical", command=self.queue_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.queue_listbox.config(yscrollcommand=scrollbar.set)

        self.btn_add_to_queue = ttk.Button(export_panel_frame, text="Add Current View to Queue", command=self.add_to_queue, state=tk.DISABLED)
        self.btn_add_to_queue.pack(fill=tk.X, pady=5, ipady=5)
        
        self.btn_clear_queue = ttk.Button(export_panel_frame, text="Clear Queue", command=self.clear_queue, state=tk.DISABLED)
        self.btn_clear_queue.pack(fill=tk.X, pady=5)

        self.btn_export_all = ttk.Button(export_panel_frame, text="Export All to Excel", command=self.export_all_to_excel, state=tk.DISABLED)
        self.btn_export_all.pack(fill=tk.X, pady=10, ipady=8)

    # NEW: Function to generate the template file
    def get_template(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
            title="Save Template As"
        )
        if not file_path: return

        try:
            # Headers based on the user's latest image
            headers = [
                'ID', 'Too cheap', 'Cheap', 'Expensive', 'Too expensive', 
                'Filter1', 'Filter2', 'Filter3'
            ]
            template_df = pd.DataFrame(columns=headers)
            template_df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Template file saved successfully to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save template file: {e}")

    def load_excel(self):
        file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=(("Excel files", "*.xlsx *.xls"),))
        if not file_path: return

        try:
            temp_df = pd.read_excel(file_path)
            
            # UPDATED: Column mapping now matches the new template (Too cheap, then Cheap)
            psm_col_names = {
                temp_df.columns[1]: 'Too cheap',
                temp_df.columns[2]: 'Cheap',
                temp_df.columns[3]: 'Expensive',
                temp_df.columns[4]: 'Too expensive'
            }
            temp_df.rename(columns=psm_col_names, inplace=True)
            
            psm_cols = ['Too cheap', 'Cheap', 'Expensive', 'Too expensive']

            # --- NEW: Data Validation Logic ---
            original_row_count = len(temp_df)
            
            # Convert to numeric, invalid data becomes NaN
            for col in psm_cols:
                temp_df[col] = pd.to_numeric(temp_df[col], errors='coerce')

            # Drop rows with any missing values in the critical PSM columns
            temp_df.dropna(subset=psm_cols, inplace=True)
            
            # Create a boolean mask for the logical check
            valid_mask = (temp_df['Too cheap'] <= temp_df['Cheap']) & \
                         (temp_df['Cheap'] <= temp_df['Expensive']) & \
                         (temp_df['Expensive'] <= temp_df['Too expensive'])
            
            self.df = temp_df[valid_mask].copy()
            validated_row_count = len(self.df)
            
            rows_dropped = original_row_count - validated_row_count
            if rows_dropped > 0:
                messagebox.showwarning(
                    "Data Validation",
                    f"{rows_dropped} rows were removed due to missing data or incorrect logic "
                    f"(e.g., 'Cheap' price > 'Expensive' price).\n\n"
                    f"Proceeding with {validated_row_count} valid rows."
                )

            if self.df.empty:
                messagebox.showerror("Error", "No valid data rows found after validation. Please check your file.")
                return
            # --- End of Validation Logic ---

            self.file_label.config(text=file_path.split('/')[-1])
            self.filtered_df = self.df.copy()
            self.create_filter_widgets()
            self.run_psm_analysis()
            self.clear_queue() 
            self.btn_add_to_queue.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error Loading File", f"An error occurred: {e}")

    def create_filter_widgets(self):
        for widget in self.filter_widgets: widget.destroy()
        self.filter_widgets.clear()
        if self.df.shape[1] > 5:
            filter_cols = self.df.columns[5:]
            for i, col in enumerate(filter_cols):
                lbl = ttk.Label(self.filter_frame, text=f"{col}:")
                lbl.grid(row=0, column=i*2, padx=(10, 2), pady=2, sticky='w')
                options = ['All'] + sorted(self.df[col].dropna().unique().tolist())
                var = tk.StringVar(value='All')
                dropdown = ttk.Combobox(self.filter_frame, textvariable=var, values=options, state='readonly', width=15)
                dropdown.grid(row=0, column=i*2 + 1, padx=(0, 10), pady=2, sticky='w')
                self.filter_widgets.extend([lbl, dropdown])
            
            btn_apply = ttk.Button(self.filter_frame, text="Apply Filters", command=self.apply_filters)
            btn_apply.grid(row=1, column=0, columnspan=2, pady=5)
            btn_reset = ttk.Button(self.filter_frame, text="Reset Filters", command=self.reset_filters)
            btn_reset.grid(row=1, column=2, columnspan=2, pady=5, padx=5)
            self.filter_widgets.extend([btn_apply, btn_reset])
            
    def apply_filters(self):
        if self.df is None: return
        temp_df = self.df.copy()
        filter_cols = self.df.columns[5:]
        for i, col in enumerate(filter_cols):
            dropdown = self.filter_widgets[i*2 + 1]
            value = dropdown.get()
            if value != 'All':
                if pd.api.types.is_numeric_dtype(temp_df[col].dropna()):
                    try:
                        value_type = temp_df[col].dropna().iloc[0].__class__
                        value = value_type(value)
                    except (ValueError, TypeError): pass
                temp_df = temp_df[temp_df[col] == value]
        self.filtered_df = temp_df
        self.run_psm_analysis()

    def reset_filters(self):
        for widget in self.filter_widgets:
            if isinstance(widget, ttk.Combobox): widget.set('All')
        if self.df is not None:
            self.filtered_df = self.df.copy()
            self.run_psm_analysis()

    def run_psm_analysis(self):
        self.current_intersections = {} 
        if self.filtered_df is None or self.filtered_df.empty or len(self.filtered_df) < 2:
            self.ax.clear()
            self.ax.text(0.5, 0.5, 'No data to display.', horizontalalignment='center', verticalalignment='center')
            self.canvas.draw()
            self.psm_data, self.current_fig = None, None
            return

        valid_df = self.filtered_df # Data is already validated
        psm_cols = ['Too cheap', 'Cheap', 'Expensive', 'Too expensive']
            
        prices = np.unique(valid_df[psm_cols].values.flatten())
        prices = prices[~np.isnan(prices)]
        
        expensive = [np.mean(valid_df['Expensive'] <= p) for p in prices]
        too_expensive = [np.mean(valid_df['Too expensive'] <= p) for p in prices]
        cheap = [np.mean(valid_df['Cheap'] >= p) for p in prices]
        too_cheap = [np.mean(valid_df['Too cheap'] >= p) for p in prices]

        self.psm_data = pd.DataFrame({'Price': prices, 'Too cheap %': too_cheap, 'Cheap %': cheap, 'Expensive %': expensive, 'Too expensive %': too_expensive})
        
        def find_intersection(x, y1, y2):
            diff, idx = np.array(y1) - np.array(y2), np.where(np.diff(np.sign(np.array(y1) - np.array(y2))))[0]
            if len(idx) == 0: return None
            idx = idx[0]
            x1, x2, y1_1, y1_2, y2_1, y2_2 = x[idx], x[idx+1], y1[idx], y1[idx+1], y2[idx], y2[idx+1]
            m1 = (y1_2 - y1_1) / (x2 - x1) if (x2 - x1) != 0 else 0
            m2 = (y2_2 - y2_1) / (x2 - x1) if (x2 - x1) != 0 else 0
            if m1 == m2: return (x1 + x2) / 2
            c1, c2 = y1_1 - m1 * x1, y2_1 - m2 * x2
            return (c2 - c1) / (m1 - m2)

        intersections = {'OPP': find_intersection(prices, too_cheap, too_expensive), 'IDP': find_intersection(prices, cheap, expensive), 'PMC': find_intersection(prices, too_cheap, expensive), 'PME': find_intersection(prices, cheap, too_expensive)}
        self.current_intersections = intersections
        self.plot_psm(intersections, len(valid_df))
    
    def plot_psm(self, intersections, n_size):
        self.ax.clear()
        sns.set_style("whitegrid")
        self.ax.plot(self.psm_data['Price'], self.psm_data['Too cheap %'], label='Too cheap')
        self.ax.plot(self.psm_data['Price'], self.psm_data['Cheap %'], label='Cheap')
        self.ax.plot(self.psm_data['Price'], self.psm_data['Expensive %'], label='Expensive')
        self.ax.plot(self.psm_data['Price'], self.psm_data['Too expensive %'], label='Too expensive')
        
        colors = {'PMC': 'green', 'PME': 'red', 'IDP': 'purple', 'OPP': 'blue'}
        labels = [f'{name}: {price:,.2f}' for name, price in intersections.items() if price is not None]
        for name, price in intersections.items():
            if price is not None: self.ax.axvline(x=price, color=colors[name], linestyle='--', lw=1.5)
        
        if labels: self.ax.text(0.98, 0.98, '\n'.join(labels), transform=self.ax.transAxes, verticalalignment='top', horizontalalignment='right', bbox=dict(boxstyle='round,pad=0.5', fc='white', alpha=0.7))

        self.ax.set_title(f'Price Sensitivity Meter (n={n_size})')
        self.ax.set_xlabel('Price')
        self.ax.set_ylabel('Percentage of Respondents')
        self.ax.yaxis.set_major_formatter(plt.FuncFormatter('{:.0%}'.format))
        self.ax.legend(loc='best')
        self.fig.tight_layout()
        self.canvas.draw()
        self.current_fig = self.fig

    def add_to_queue(self):
        if self.psm_data is None or self.current_fig is None:
            messagebox.showwarning("Warning", "No valid data/chart to add.")
            return

        filter_name_parts = [f"{col}-{str(self.filter_widgets[i*2 + 1].get()).strip()}" for i, col in enumerate(self.df.columns[5:]) if str(self.filter_widgets[i*2 + 1].get()).strip() != 'All']
        queue_name = "_".join(filter_name_parts) if filter_name_parts else "Total"
        
        if queue_name in self.export_queue and not messagebox.askyesno("Confirm", f"'{queue_name}' is already in the queue. Overwrite?"):
            return
        
        self.export_queue[queue_name] = {
            'data': self.psm_data.copy(), 
            'figure': copy.deepcopy(self.current_fig), 
            'intersections': self.current_intersections.copy()
        }
        self.update_queue_listbox()
        
    def update_queue_listbox(self):
        self.queue_listbox.delete(0, tk.END)
        for name in self.export_queue.keys(): self.queue_listbox.insert(tk.END, name)
        
        state = tk.NORMAL if self.export_queue else tk.DISABLED
        self.btn_export_all.config(state=state)
        self.btn_clear_queue.config(state=state)
            
    def clear_queue(self):
        self.export_queue.clear()
        self.update_queue_listbox()
        
    def export_all_to_excel(self):
        if not self.export_queue:
            messagebox.showinfo("Info", "Export queue is empty.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")], title="Save Export As")
        if not file_path: return

        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                workbook = writer.book
                header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1})
                default_format = workbook.add_format({'num_format': '0.00'})
                metric_formats = {
                    'OPP': workbook.add_format({'num_format': '0.00', 'font_color': 'blue', 'bold': True}),
                    'IDP': workbook.add_format({'num_format': '0.00', 'font_color': 'purple', 'bold': True}),
                    'PMC': workbook.add_format({'num_format': '0.00', 'font_color': 'green', 'bold': True}),
                    'PME': workbook.add_format({'num_format': '0.00', 'font_color': 'red', 'bold': True})
                }

                for sheet_name, content in self.export_queue.items():
                    safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '_', '-')).rstrip()[:31]
                    df_to_write = content['data']
                    df_to_write.to_excel(writer, sheet_name=safe_sheet_name, index=False, startrow=1)
                    
                    worksheet = writer.sheets[safe_sheet_name]
                    worksheet.write(0, 0, f"PSM Analysis: {safe_sheet_name}", workbook.add_format({'bold': True, 'font_size': 14}))

                    start_row = len(df_to_write) + 4 
                    worksheet.merge_range(start_row, 0, start_row, 1, 'Key Metrics', header_format)
                    
                    metrics_data = content['intersections']
                    row = start_row + 1
                    for key, value in metrics_data.items():
                        worksheet.write(row, 0, key)
                        cell_format = metric_formats.get(key, default_format)
                        if value is not None:
                            worksheet.write(row, 1, value, cell_format)
                        else:
                            worksheet.write(row, 1, "Not Found")
                        row += 1
                    
                    img_buffer = io.BytesIO()
                    fig = content['figure']
                    fig.savefig(img_buffer, format='png', dpi=200, bbox_inches='tight')
                    img_buffer.seek(0)
                    
                    worksheet.insert_image('G2', 'psm_chart.png', {'image_data': img_buffer, 'x_scale': 0.8, 'y_scale': 0.8})

            messagebox.showinfo("Success", f"Successfully exported all items to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred during export: {e}")



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
        app = PSMApp(root)
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
