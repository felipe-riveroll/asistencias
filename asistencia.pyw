import pandas as pd
import os
import datetime
import subprocess
import traceback # Keep for debugging if needed
from tkinter import Tk, Label, Button, Entry, StringVar, filedialog, Frame, ttk, messagebox, Toplevel
from PIL import Image, ImageTk
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side

pd.options.mode.chained_assignment = 'raise'

class CheckadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Checadas")
        self.root.geometry("750x520")
        self.root.resizable(True, True)

        # ────────────  Colores / estilos  ────────────
        self.primary_color   = "#2c3e50"
        self.secondary_color = "#3498db"
        self.bg_color        = "#f5f5f5"
        self.text_color      = "#333333"
        self.success_color   = "#27ae60"
        self.warning_color   = "#f39c12"
        self.error_color     = "#e74c3c"
        self.root.configure(bg=self.bg_color)

        # Cargar datos de horas esperadas
        self.expected_hours_df = self._load_expected_hours_data()
        # Caché para valores de horas esperadas (mejora de rendimiento)
        self.expected_hours_cache = {}

        # ────────────  Barra de estado  ────────────
        self.status_frame = Frame(root, bg="#e0e0e0", relief="ridge", bd=1)
        self.status_frame.pack(side="bottom", fill="x")
        self.status_label = Label(self.status_frame, text="Listo para procesar", font=("Segoe UI", 10),
                                  bg="#e0e0e0", fg=self.primary_color, padx=10, pady=8)
        self.status_label.pack(fill="x")

        # ttk base
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TEntry",  font=("Segoe UI", 10))
        style.configure("TLabel",  font=("Segoe UI", 10), background=self.bg_color)

        # ────────────  Vars  ────────────
        self.input_file_path = StringVar()
        self.output_file_name = StringVar(value=f"reporte_checador_{datetime.datetime.now().strftime('%d%m%Y')}")

        # ────────────  Layout  ────────────
        main = Frame(root, bg=self.bg_color, padx=30, pady=20); main.pack(fill="both", expand=True)

        # logo
        try:
            # Ensure "Logo_asia.png" is in the same directory as the script or provide a full path
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logo_asia.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path); w, h = img.size; nw = 200; nh = int(nw / w * h)
                logo = ImageTk.PhotoImage(img.resize((nw, nh), Image.LANCZOS))
                Label(main, image=logo, bg=self.bg_color).pack(pady=(0, 15)); self._logo_ref = logo
        except Exception as e_logo:
            print(f"Error loading logo: {e_logo}")
            pass

        # banner
        banner = Frame(main, bg=self.primary_color, height=60); banner.pack(fill="x", pady=(0, 20))
        Label(banner, text="Procesador de Checadas de Empleados", font=("Segoe UI", 18, "bold"), bg=self.primary_color, fg="white").pack(pady=10)

        form = Frame(main, bg=self.bg_color, padx=20, pady=10); form.pack(fill="both", expand=True)
        row1 = Frame(form, bg=self.bg_color, pady=10); row1.pack(fill="x")
        Label(row1, text="Archivo Excel de Checadas:", font=("Segoe UI", 11), bg=self.bg_color, fg=self.text_color).pack(side="left", padx=(0, 10))
        Entry(row1, textvariable=self.input_file_path, font=("Segoe UI", 10), bd=1, relief="solid").pack(side="left", fill="x", expand=True, ipady=3)
        Button(row1, text="Examinar...", command=self.browse_file, font=("Segoe UI", 10), bg=self.secondary_color, fg="white", relief="flat").pack(side="left", padx=(10, 0))

        row2 = Frame(form, bg=self.bg_color, pady=10); row2.pack(fill="x")
        Label(row2, text="Nombre del archivo de salida:", font=("Segoe UI", 11), bg=self.bg_color, fg=self.text_color).pack(side="left", padx=(0, 10))
        Entry(row2, textvariable=self.output_file_name, font=("Segoe UI", 10), bd=1, relief="solid").pack(side="left", fill="x", expand=True, ipady=3)
        Label(row2, text=".xlsx", font=("Segoe UI", 11), bg=self.bg_color, fg=self.text_color).pack(side="left")

        ttk.Separator(form, orient="horizontal").pack(fill="x", pady=20)

        actions = Frame(form, bg=self.bg_color, pady=20); actions.pack(fill="x")
        self.process_button = Button(actions, text="Procesar Archivo", command=self.generate_report, font=("Segoe UI", 12, "bold"), bg=self.secondary_color, fg="white", relief="flat", padx=20, pady=8)
        self.process_button.pack(pady=10)
        self.progress = ttk.Progressbar(actions, orient="horizontal", length=500, mode="indeterminate")

    # ────────────  Cargar horas esperadas  ────────────
    def _load_expected_hours_data(self):
        """
        Carga el archivo CSV con las horas esperadas por empleado y día.
        Expected CSV format: Employee (ID), Lunes (seconds), Martes (seconds), ...
        """
        try:
            base_path = os.path.dirname(os.path.abspath(__file__))
            expected_hours_path = os.path.join(base_path, 'expected_hours_data.csv')
            
            if os.path.exists(expected_hours_path):
                with open(expected_hours_path, 'r', encoding='utf-8') as f:
                    first_line = f.readline().strip()
                    skip_rows = 1 if first_line.startswith('//') else 0
                
                df_expected = pd.read_csv(expected_hours_path, skiprows=skip_rows)
                
                required_cols = ['Employee', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
                if all(col in df_expected.columns for col in required_cols):
                    df_expected['Employee'] = pd.to_numeric(df_expected['Employee'], errors='coerce').fillna(0).astype(int)
                    for day_col in ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']:
                        if day_col in df_expected.columns:
                            # Ensure these columns are numeric (expected to be seconds)
                            df_expected[day_col] = pd.to_numeric(df_expected[day_col], errors='coerce').fillna(0)
                    print(f"Datos de horas esperadas cargados: {len(df_expected)} empleados. Columnas: {df_expected.columns.tolist()}")
                    # print(df_expected.head()) # Optional: print head for debugging
                    return df_expected
                else:
                    print(f"Advertencia: Faltan columnas requeridas en el archivo de horas esperadas. Se esperaban: {required_cols}. Se encontraron: {df_expected.columns.tolist()}")
            else:
                messagebox.showwarning("Archivo no encontrado", f"No se encontró el archivo de horas esperadas en {expected_hours_path}. La columna 'Horas esperadas' será 0.")
                print(f"Advertencia: No se encontró el archivo de horas esperadas en {expected_hours_path}")
        except Exception as e:
            messagebox.showerror("Error al cargar horas", f"Error al cargar horas esperadas: {e}\n{traceback.format_exc()}")
            print(f"Error al cargar horas esperadas: {e}\n{traceback.format_exc()}")
        return None

    # ────────────  Utilidades de UI  ────────────
    def browse_file(self):
        fp = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx;*.xls"), ("Todos", "*.*")])
        if fp:
            self.input_file_path.set(fp)
            self.status_label.configure(text=f"Archivo seleccionado: {os.path.basename(fp)}", fg=self.primary_color, bg="#e0e0e0")

    def _set_status(self, msg, kind="info"):
        cfg = {"success":("white", self.success_color), "warning":("white", self.warning_color), "error":("white", self.error_color), "info":(self.primary_color, "#e0e0e0")}
        fg, bg = cfg.get(kind, cfg["info"]); self.status_label.configure(text=msg, fg=fg, bg=bg); self.root.update()

    def _toggle_busy(self, busy: bool):
        if busy:
            self.process_button.configure(state="disabled", text="Procesando...", bg="#95a5a6")
            self.progress.pack(pady=10); self.progress.start(10)
        else:
            self.process_button.configure(state="normal", text="Procesar Archivo", bg=self.secondary_color)
            self.progress.stop(); self.progress.pack_forget()

    # ────────────  Diálogo de éxito con botón abrir  ────────────
    def _show_success_dialog(self, path: str):
        dlg = Toplevel(self.root); dlg.title("Proceso completado"); dlg.geometry("450x250"); dlg.resizable(False, False); dlg.configure(bg="white")
        content = Frame(dlg, bg="white", padx=20, pady=10); content.pack(fill="both", expand=True)
        try:
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logo_asia.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path); w, h = img.size; nw = 150; nh = int(nw / w * h)
                logo = ImageTk.PhotoImage(img.resize((nw, nh), Image.LANCZOS)); Label(content, image=logo, bg="white").pack(pady=(0,10)); dlg._logo=logo
            else:
                 Label(content, text="ⓘ", font=("Segoe UI", 24), fg=self.secondary_color, bg="white").pack(pady=(10,5))
        except Exception:
            Label(content, text="ⓘ", font=("Segoe UI", 24), fg=self.secondary_color, bg="white").pack(pady=(10,5))
        Label(content, text="El reporte se ha generado correctamente.", font=("Segoe UI", 11), bg="white").pack(pady=5)
        short = path if len(path)<=50 else os.path.join(os.path.dirname(path)[:10]+"...", os.path.basename(path))
        pf = Frame(content, bg="white"); pf.pack(pady=5, fill="x")
        Label(pf, text="Ubicación:", font=("Segoe UI", 10), bg="white").pack(side="left")
        Label(pf, text=short, font=("Segoe UI", 9), fg="#555", bg="white").pack(side="left")
        bf = Frame(content, bg="white"); bf.pack(fill="x", pady=10)
        def _open_file_action(): # Renamed to avoid conflict with builtin open
            try:
                if os.name=='nt': os.startfile(path)
                elif os.name=='posix': subprocess.call(('open' if os.path.exists('/usr/bin/open') else 'xdg-open', path))
                dlg.destroy()
            except Exception as e_open:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e_open}")
        Button(bf, text="Abrir archivo", font=("Segoe UI", 10), bg=self.secondary_color, fg="white", width=12, command=_open_file_action).pack(side="right", padx=5)
        Button(bf, text="Aceptar", font=("Segoe UI", 10), bg="#f0f0f0", width=10, command=dlg.destroy).pack(side="right", padx=5)
        dlg.transient(self.root); dlg.grab_set(); dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.update_idletasks(); w_dlg, h_dlg = dlg.winfo_width(), dlg.winfo_height(); x = (dlg.winfo_screenwidth()-w_dlg)//2; y = (dlg.winfo_screenheight()-h_dlg)//2; dlg.geometry(f"{w_dlg}x{h_dlg}+{x}+{y}")

    # ────────────  Lógica principal  ────────────
    def generate_report(self):
        src = self.input_file_path.get().strip()
        if not src:
            self._set_status("Seleccione un archivo", "error"); messagebox.showerror("Error", "Debe seleccionar un archivo de entrada"); return
        
        output_filename = self.output_file_name.get().strip() + ".xlsx"
        if os.path.isabs(src) and os.path.isdir(os.path.dirname(src)):
            dst_folder = os.path.dirname(src)
        else: 
            dst_folder = os.path.dirname(os.path.abspath(__file__))
        dst = os.path.join(dst_folder, output_filename)

        try:
            self._toggle_busy(True); self._set_status("Procesando archivo...", "info")
            df = pd.read_excel(src) # Original DataFrame from Excel
            if {'Employee Name','Time'}.difference(df.columns):
                raise ValueError("Las columnas requeridas 'Employee Name' y 'Time' no se encontraron.")
            
            # --- Initial Data Preparation ---
            df_processed = df.copy() # Work on a copy
            df_processed['Time'] = pd.to_datetime(df_processed['Time'], errors='coerce')
            df_processed.dropna(subset=['Time'], inplace=True) # Remove rows where Time could not be parsed

            df_processed['Day_raw'] = df_processed['Time'].dt.date
            df_processed['WorkDay'] = df_processed.apply(lambda r: r['Day_raw'] - datetime.timedelta(days=1) if r['Time'].hour < 6 else r['Day_raw'], axis=1)
            if 'Shift' not in df_processed.columns: df_processed['Shift'] = ''
            df_processed['Shift'] = df_processed['Shift'].fillna('')

            # Separate records with shift and without
            df_turno = df_processed[df_processed['Shift'] != ''].copy()
            df_sin_turno = df_processed[df_processed['Shift'] == ''].copy()
            df_sin_turno['Merged'] = False

            # Group records with shift
            grouped = (df_turno.groupby(['Employee Name', 'Shift', 'WorkDay'])
                       .agg(checadas_list=('Time', lambda ts: sorted([pd.to_datetime(t) for t in ts if pd.notnull(t)])))
                       .reset_index())

            # --- Merge records without shift into existing groups ---
            def merge_no_shift(gdf, ns_df):
                res_df = gdf.copy()
                if 'checadas_list' not in res_df.columns: # Should exist from previous agg
                    res_df['checadas_list'] = [[] for _ in range(len(res_df))]
                else: # Ensure all are lists
                    res_df['checadas_list'] = res_df['checadas_list'].apply(lambda x: list(x) if isinstance(x, (list, pd.Series)) else [])

                for idx, row_ns in ns_df.iterrows():
                    time_to_add = pd.to_datetime(row_ns['Time'])
                    mask = (res_df['Employee Name'] == row_ns['Employee Name']) & (res_df['WorkDay'] == row_ns['WorkDay'])
                    if mask.any():
                        res_idx = res_df[mask].index[0]
                        current_list = res_df.at[res_idx, 'checadas_list']
                        if time_to_add not in current_list:
                            res_df.at[res_idx, 'checadas_list'] = sorted(current_list + [time_to_add])
                        ns_df.loc[idx, 'Merged'] = True # Mark as merged
                    else:
                        ns_df.loc[idx, 'Merged'] = False


                extra_df = ns_df[~ns_df['Merged']]
                if not extra_df.empty:
                    add_df = (extra_df.groupby(['Employee Name', 'WorkDay'])
                              .agg(checadas_list=('Time', lambda ts: sorted([pd.to_datetime(t) for t in ts if pd.notnull(t)])))
                              .reset_index())
                    add_df['Shift'] = '' # No shift for these new groups
                    res_df = pd.concat([res_df, add_df], ignore_index=True)
                return res_df

            grouped = merge_no_shift(grouped, df_sin_turno)
            grouped.rename(columns={'WorkDay': 'Fecha_raw'}, inplace=True) # This is the actual date of work
            grouped.sort_values(['Employee Name', 'Fecha_raw'], inplace=True)

            # --- Calculate total worked hours (Timedelta) and format as HH:MM:SS string ---
            def calc_actual_worked_hours(lst_times):
                if not lst_times or len(lst_times) < 2: return pd.Timedelta(0)
                # Already ensured lst_times are sorted datetime objects
                return lst_times[-1] - lst_times[0]

            grouped['total_timedelta_actual'] = grouped['checadas_list'].apply(calc_actual_worked_hours)
            
            fmt_timedelta_to_str = lambda td: f"{int(td.total_seconds()//3600):02d}:{int(td.total_seconds()%3600//60):02d}:{int(round(td.total_seconds()%60)):02d}" if pd.notnull(td) and td.total_seconds() > 0 else "00:00:00"
            grouped['Horas totales_str'] = grouped['total_timedelta_actual'].apply(fmt_timedelta_to_str)

            # --- Prepare Checada columns ---
            grouped['Checadas_str_list'] = grouped['checadas_list'].apply(lambda ts: [t.strftime('%H:%M:%S') for t in ts if pd.notnull(t)])
            max_chec = grouped['Checadas_str_list'].str.len().max()
            if pd.isna(max_chec) or max_chec == 0: max_chec = 1
            
            chec_df_data = {}
            for i in range(int(max_chec)):
                chec_df_data[f'Checada {i+1}'] = grouped['Checadas_str_list'].apply(lambda x: x[i] if i < len(x) else None)
            chec_df = pd.DataFrame(chec_df_data, index=grouped.index) # Align index for concat

            # --- Employee ID Mapping (from original df to grouped df) ---
            # Using 'Employee Name' as key. Ensure 'Employee Name' exists in 'df' (original excel)
            # And 'Employee' (ID column) exists in 'df'
            if 'Employee' in df.columns and 'Employee Name' in df.columns:
                # Create map from original df, ensure Employee Name is unique or take first ID
                id_map = df[['Employee Name', 'Employee']].drop_duplicates('Employee Name').set_index('Employee Name')['Employee'].to_dict()
                grouped['ID Empleado_val'] = grouped['Employee Name'].map(id_map)
            else:
                # Fallback: try to infer if 'Employee' column is missing but ID might be in another column
                # This part depends heavily on the structure of the input Excel if 'Employee' column is absent
                # For now, if 'Employee' column is not there, ID will be blank or from a prior assumption
                grouped['ID Empleado_val'] = grouped.get('Employee', "") # .get if 'Employee' was from a previous step

            # --- Create the main report DataFrame ('Detalle' sheet) ---
            report_cols_from_grouped = ['ID Empleado_val', 'Employee Name', 'Shift', 'Fecha_raw', 'Horas totales_str']
            # Ensure columns exist in grouped
            for col_name in report_cols_from_grouped:
                if col_name not in grouped:
                    grouped[col_name] = None if col_name != 'Shift' else ''


            report_df = pd.concat([grouped[report_cols_from_grouped], chec_df], axis=1)
            report_df.rename(columns={'ID Empleado_val': 'ID Empleado',
                                   'Employee Name': 'Nombre del empleado',
                                   'Shift': 'Turno',
                                   'Fecha_raw': 'Fecha', # This is the actual date for the row
                                   'Horas totales_str': 'Horas totales'}, inplace=True)


            # --- Add Spanish Day Name ---
            report_df['Fecha'] = pd.to_datetime(report_df['Fecha'], errors='coerce')
            dias_semana = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves', 4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}
            report_df['Día'] = report_df['Fecha'].apply(lambda x: dias_semana[x.weekday()] if pd.notnull(x) and hasattr(x, 'weekday') else '')

            # --- Add "Horas esperadas" (as numeric seconds) ---
            def get_expected_seconds_for_day(row_data):
                if self.expected_hours_df is None: return 0.0 # No expected hours data loaded
                try:
                    # 'Turno' in row_data comes from 'report_df'
                    if row_data.get('Turno') == 'Totales': return 0.0 # Don't calculate for summary rows yet
                    
                    # 'Fecha' and 'Día' (Spanish name) from row_data
                    if not pd.notnull(row_data.get('Fecha')) or not hasattr(row_data['Fecha'], 'weekday'): return 0.0
                    
                    dia_semana_str = row_data.get('Día', '')
                    if not dia_semana_str or dia_semana_str not in self.expected_hours_df.columns:
                        return 0.0
                    
                    emp_id_str = str(row_data.get('ID Empleado', "")).strip()
                    if not emp_id_str: return 0.0
                    
                    cache_key = f"{emp_id_str}_{dia_semana_str}"
                    if cache_key in self.expected_hours_cache:
                        return float(self.expected_hours_cache[cache_key])
                    
                    try: # Convert emp_id_str to numeric for matching with 'Employee' column
                        employee_id_num = int(float(emp_id_str)) 
                    except ValueError:
                        return 0.0 # Cannot convert to number for matching
                        
                    emp_mask = self.expected_hours_df['Employee'] == employee_id_num
                    if not emp_mask.any():
                        self.expected_hours_cache[cache_key] = 0.0
                        return 0.0
                    
                    # Get value (seconds) from the CSV data for that employee and day
                    value_seconds = self.expected_hours_df.loc[emp_mask, dia_semana_str].iloc[0] # use iloc[0] for Series
                    result_seconds = float(value_seconds) if pd.notnull(value_seconds) else 0.0
                    
                    self.expected_hours_cache[cache_key] = result_seconds
                    return result_seconds
                except Exception as e_get_exp:
                    # print(f"Error en get_expected_seconds_for_day para ID {row_data.get('ID Empleado', 'N/A')}, Día {dia_semana_str}: {e_get_exp}")
                    return 0.0
            
            report_df['Horas esperadas'] = report_df.apply(get_expected_seconds_for_day, axis=1) # Numeric seconds

            # --- Reorder columns for the "Detalle" sheet ---
            core_cols = ['ID Empleado', 'Nombre del empleado', 'Turno', 'Fecha', 'Día', 'Horas esperadas', 'Horas totales']
            checada_cols_in_report = sorted([col for col in report_df.columns if col.startswith('Checada ')], 
                                            key=lambda x: int(x.split(' ')[1]))
            final_report_columns_ordered = core_cols + checada_cols_in_report
            
            # Ensure all final columns exist in report_df, add if missing (though they should be there)
            for col_name in final_report_columns_ordered:
                if col_name not in report_df.columns:
                    report_df[col_name] = None 
            report_df = report_df[final_report_columns_ordered]


            # --- Summary DataFrame ("Resumen" sheet) ---
            # Use 'grouped' for this summary as it has 'total_timedelta_actual'
            summary_group_by_cols = ['ID Empleado_val', 'Employee Name'] \
                if 'ID Empleado_val' in grouped.columns and grouped['ID Empleado_val'].notna().any() \
                else ['Employee Name']

            summary = (grouped.groupby(summary_group_by_cols)
                       .agg(first_day_worked=('Fecha_raw', 'min'),
                            last_day_worked=('Fecha_raw', 'max'),
                            days_actually_worked=('Fecha_raw', 'nunique'),
                            total_worked_time_sum=('total_timedelta_actual', 'sum'))
                       .reset_index())

            if 'ID Empleado_val' not in summary.columns: summary['ID Empleado_val'] = ""
            
            summary.rename(columns={'ID Empleado_val': 'ID Empleado', 'Employee Name': 'Nombre'}, inplace=True)
            summary['Días del periodo'] = (summary['last_day_worked'] - summary['first_day_worked']).apply(lambda td: td.days + 1 if pd.notnull(td) else 0)
            summary['Horas trabajadas_str_sum'] = summary['total_worked_time_sum'].apply(fmt_timedelta_to_str)
            
            resumen_df_cols = ['ID Empleado', 'Nombre', 'Días del periodo', 'days_actually_worked', 'Horas trabajadas_str_sum']
            resumen_df = summary[resumen_df_cols].rename(columns={
                'days_actually_worked': 'Días trabajados',
                'Horas trabajadas_str_sum': 'Horas trabajadas'})


            # --- Create "Totales" rows to append to "Detalle" sheet ---
            total_rows_for_detail_list = []
            for _, r_summary_row in resumen_df.iterrows():
                emp_name_for_total = r_summary_row['Nombre']
                
                # Sum numeric 'Horas esperadas' from daily records for this employee
                # report_df contains daily details, including numeric 'Horas esperadas'
                current_emp_daily_details = report_df[
                    (report_df['Nombre del empleado'] == emp_name_for_total) &
                    (report_df['Turno'] != 'Totales') # Exclude any pre-existing summary rows
                ]
                sum_numeric_expected_seconds_for_emp = pd.to_numeric(current_emp_daily_details['Horas esperadas'], errors='coerce').sum()

                total_row_dict = {
                    'ID Empleado': r_summary_row.get('ID Empleado', ""),
                    'Nombre del empleado': emp_name_for_total,
                    'Turno': 'Totales',
                    'Fecha': int(r_summary_row['Días trabajados']) if pd.notnull(r_summary_row['Días trabajados']) else 0,
                    'Día': '', 
                    'Horas esperadas': sum_numeric_expected_seconds_for_emp, # Sum of numeric daily expected seconds
                    'Horas totales': r_summary_row['Horas trabajadas'] # HH:MM:SS sum of actual worked time from resumen_df
                }
                # Add blank checada columns for the total row
                for c_col in checada_cols_in_report:
                    total_row_dict[c_col] = ''
                total_rows_for_detail_list.append(total_row_dict)
            
            totals_to_append_df = pd.DataFrame(total_rows_for_detail_list)
            if not totals_to_append_df.empty:
                totals_to_append_df = totals_to_append_df.reindex(columns=report_df.columns)


            # --- Combine daily details with their "Totales" row in order ---
            final_detail_sheet_dfs = []
            if not report_df.empty:
                report_df['Fecha'] = pd.to_datetime(report_df['Fecha'], errors='coerce') # Ensure datetime for sort
                for emp_name, daily_data_group in report_df.groupby('Nombre del empleado', sort=False):
                    final_detail_sheet_dfs.append(daily_data_group.sort_values('Fecha'))
                    
                    emp_total_row = totals_to_append_df[totals_to_append_df['Nombre del empleado'] == emp_name]
                    if not emp_total_row.empty:
                        final_detail_sheet_dfs.append(emp_total_row)
            
            if not final_detail_sheet_dfs:
                final_detail_report_df = pd.DataFrame(columns=final_report_columns_ordered)
            else:
                final_detail_report_df = pd.concat(final_detail_sheet_dfs, ignore_index=True)

            # --- Write to Excel ---
            with pd.ExcelWriter(dst, engine='openpyxl') as writer:
                final_detail_report_df.to_excel(writer, index=False, sheet_name='Detalle')
                resumen_df.to_excel(writer, index=False, sheet_name='Resumen')

            self._format_excel(dst)
            self._toggle_busy(False); self._set_status("Reporte generado exitosamente", "success")
            self._show_success_dialog(dst)

        except ValueError as ve:
            self._toggle_busy(False); self._set_status(f"Error de valor: {ve}", "error"); 
            messagebox.showerror("Error de Valor", f"Ocurrió un error con los datos:\n{ve}\n\n{traceback.format_exc()}")
            print(traceback.format_exc())
        except Exception as e:
            self._toggle_busy(False); self._set_status(f"Error: {e}", "error"); 
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{e}\n\n{traceback.format_exc()}")
            print(traceback.format_exc())

    # ────────────  Formato Excel  ────────────
    def _format_excel(self, path):
        wb=load_workbook(path)
        def _format_ws(ws):
            header_fill=PatternFill(start_color="3498DB", end_color="3498DB", fill_type="solid")
            header_font=Font(color="FFFFFF", bold=True)
            total_fill=PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Light gray for totals
            bold_font=Font(bold=True)
            thin=Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            if ws.max_row == 0: return

            # Format header
            for c_idx_plus_1 in range(1, ws.max_column + 1):
                cell = ws.cell(1, c_idx_plus_1)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin
            
            # Format "Totales" rows in "Detalle" sheet
            if ws.title == 'Detalle':
                turno_col_letter = None
                # Find 'Turno' column letter dynamically
                for cell_header in ws[1]: # ws[1] is the first row (header)
                    if cell_header.value == 'Turno':
                        turno_col_letter = cell_header.column_letter
                        break
                
                if turno_col_letter:
                    for r_idx_plus_1 in range(2, ws.max_row + 1): # Start from row 2
                        # Access cell by letter and row index
                        if ws[f"{turno_col_letter}{r_idx_plus_1}"].value == 'Totales':
                            for c_idx_plus_1_total in range(1, ws.max_column + 1):
                                cell_total = ws.cell(r_idx_plus_1, c_idx_plus_1_total)
                                cell_total.fill = total_fill
                                cell_total.font = bold_font
                                cell_total.border = thin
            
            # Adjust column widths
            for col_letter_obj in ws.columns: # ws.columns gives tuples of cells for each column
                column_letter_str = col_letter_obj[0].column_letter # Get letter from first cell of column
                
                # Calculate max length for the column
                max_len = 0
                for cell_in_col in col_letter_obj:
                    if cell_in_col.value is not None:
                        try:
                            max_len = max(max_len, len(str(cell_in_col.value)))
                        except: pass # Ignore if len fails
                
                adjusted_width = max_len + 2 # Basic padding

                # Specific widths based on header name (more robust)
                header_value = ws[f"{column_letter_str}1"].value # Header value for this column

                if header_value == 'ID Empleado': adjusted_width = max(adjusted_width, 12)
                elif header_value == 'Nombre del empleado': adjusted_width = max(adjusted_width, 30)
                elif header_value == 'Nombre': adjusted_width = max(adjusted_width, 30) # For Resumen sheet
                elif header_value == 'Turno': adjusted_width = max(adjusted_width, 10)
                elif header_value == 'Fecha': adjusted_width = max(adjusted_width, 12)
                elif header_value == 'Día': adjusted_width = max(adjusted_width, 12)
                elif header_value in ['Horas esperadas', 'Horas totales', 'Horas trabajadas']: adjusted_width = max(adjusted_width, 15)
                elif str(header_value).startswith('Checada'): adjusted_width = max(adjusted_width, 10)
                elif header_value in ['Días del periodo', 'Días trabajados']: adjusted_width = max(adjusted_width, 18)
                else: adjusted_width = max(adjusted_width, 12) # Default min width
                
                ws.column_dimensions[column_letter_str].width = adjusted_width
        
        for sheet_name_iter in wb.sheetnames:
            _format_ws(wb[sheet_name_iter])
        wb.save(path)

if __name__ == "__main__":
    root=Tk(); app=CheckadorApp(root); root.mainloop()