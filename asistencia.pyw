import pandas as pd
import os
import datetime
import subprocess
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
            img = Image.open("Logo_asia.png"); w, h = img.size; nw = 200; nh = int(nw / w * h)
            logo = ImageTk.PhotoImage(img.resize((nw, nh), Image.LANCZOS))
            Label(main, image=logo, bg=self.bg_color).pack(pady=(0, 15)); self._logo_ref = logo
        except Exception:
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
            img = Image.open("Logo_asia.png"); w, h = img.size; nw = 150; nh = int(nw / w * h)
            logo = ImageTk.PhotoImage(img.resize((nw, nh), Image.LANCZOS)); Label(content, image=logo, bg="white").pack(pady=(0,10)); dlg._logo=logo
        except Exception:
            Label(content, text="ⓘ", font=("Segoe UI", 24), fg=self.secondary_color, bg="white").pack(pady=(10,5))
        Label(content, text="El reporte se ha generado correctamente.", font=("Segoe UI", 11), bg="white").pack(pady=5)

        # Ruta (acortada)
        short = path if len(path)<=50 else os.path.join(os.path.dirname(path)[:10]+"...", os.path.basename(path))
        pf = Frame(content, bg="white"); pf.pack(pady=5, fill="x")
        Label(pf, text="Ubicación:", font=("Segoe UI", 10), bg="white").pack(side="left")
        Label(pf, text=short, font=("Segoe UI", 9), fg="#555", bg="white").pack(side="left")

        # Botones
        bf = Frame(content, bg="white"); bf.pack(fill="x", pady=10)
        def _open():
            try:
                if os.name=='nt': os.startfile(path)
                elif os.name=='posix': subprocess.call(('open' if os.path.exists('/usr/bin/open') else 'xdg-open', path))
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        Button(bf, text="Abrir archivo", font=("Segoe UI", 10), bg=self.secondary_color, fg="white", width=12, command=_open).pack(side="right", padx=5)
        Button(bf, text="Aceptar", font=("Segoe UI", 10), bg="#f0f0f0", width=10, command=dlg.destroy).pack(side="right", padx=5)

        dlg.transient(self.root); dlg.grab_set(); dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.update_idletasks(); w, h = dlg.winfo_width(), dlg.winfo_height(); x = (dlg.winfo_screenwidth()-w)//2; y = (dlg.winfo_screenheight()-h)//2; dlg.geometry(f"{w}x{h}+{x}+{y}")

    # ────────────  Lógica principal  ────────────
    def generate_report(self):
        src = self.input_file_path.get().strip()
        if not src:
            self._set_status("Seleccione un archivo", "error"); messagebox.showerror("Error", "Debe seleccionar un archivo de entrada"); return
        dst = os.path.join(os.path.dirname(src), self.output_file_name.get().strip() + ".xlsx")
        try:
            self._toggle_busy(True); self._set_status("Procesando archivo...", "info")
            df = pd.read_excel(src)
            if {'Employee Name','Time'}.difference(df.columns):
                raise ValueError("Las columnas requeridas 'Employee Name' y 'Time' no se encontraron.")
            df['Time'] = pd.to_datetime(df['Time']); df['Day'] = df['Time'].dt.date
            df['WorkDay'] = df.apply(lambda r: r['Day']-datetime.timedelta(days=1) if r['Time'].hour<6 else r['Day'], axis=1)
            if 'Shift' not in df.columns: df['Shift']=''
            df['Shift'] = df['Shift'].fillna('')
            df_turno = df[df['Shift']!=''].copy(); df_sin = df[df['Shift']==''].copy(); df_sin['Merged']=False
            grouped = (df_turno.groupby(['Employee Name','Shift','WorkDay']).agg(checadas_list=('Time', lambda ts: sorted(ts))).reset_index())
            grouped['total_timedelta'] = grouped['checadas_list'].apply(lambda lst: lst[-1]-lst[0])
            def merge_no_shift(gdf, ns):
                res=gdf.copy()
                for idx,row in ns.iterrows():
                    mask=(res['Employee Name']==row['Employee Name'])&(res['WorkDay']==row['WorkDay'])
                    if mask.any():
                        i=res[mask].index[0]; lst=res.at[i,'checadas_list']
                        if row['Time'] not in lst:
                            res.at[i,'checadas_list']=sorted(lst+[row['Time']]); ns.at[idx,'Merged']=True
                extra=ns[~ns['Merged']]
                if not extra.empty:
                    add=(extra.groupby(['Employee Name','WorkDay']).agg(checadas_list=('Time', lambda ts: sorted(ts))).reset_index())
                    add['Shift']=''; res=pd.concat([res,add], ignore_index=True)
                return res
            grouped = merge_no_shift(grouped, df_sin)
            grouped.rename(columns={'WorkDay':'Day'}, inplace=True); grouped.sort_values(['Employee Name','Day'], inplace=True)
            def calc_hours(lst):
                end=lst[-1]+datetime.timedelta(days=1) if lst[-1]<lst[0] else lst[-1]
                return end-lst[0]
            grouped['total_timedelta']=grouped['checadas_list'].apply(calc_hours)
            fmt=lambda td:f"{int(td.total_seconds()//3600):02d}:{int(td.total_seconds()%3600//60):02d}:{int(td.total_seconds()%60):02d}"
            grouped['Horas totales']=grouped['total_timedelta'].apply(fmt)
            grouped['Checadas_str']=grouped['checadas_list'].apply(lambda ts:[t.strftime('%H:%M:%S') for t in ts])
            max_chec=grouped['Checadas_str'].str.len().max()
            chec_df=pd.DataFrame({f'Checada {i+1}':grouped['Checadas_str'].apply(lambda x:x[i] if i<len(x) else None) for i in range(max_chec)})
            report=pd.concat([grouped[['Employee Name','Shift','Day','Horas totales']], chec_df], axis=1)
            report.rename(columns={'Employee Name':'Nombre del empleado','Shift':'Turno','Day':'Día'}, inplace=True)
            totals=(grouped.groupby('Employee Name').agg(total_timedelta=('total_timedelta','sum'), days_worked=('Day','nunique')).reset_index())
            totals['Horas totales']=totals['total_timedelta'].apply(fmt)
            total_rows=[]
            for _,r in totals.iterrows():
                d={'Nombre del empleado':r['Employee Name'],'Turno':'Totales','Día':int(r['days_worked']),'Horas totales':r['Horas totales']}
                for c in chec_df.columns: d[c]=''
                total_rows.append(d)
            totals_df=pd.DataFrame(total_rows)
            final=[]
            for emp,grp in report.groupby('Nombre del empleado', sort=False):
                final.append(grp.sort_values('Día'))
                final.append(totals_df[totals_df['Nombre del empleado']==emp])
            final_report=pd.concat(final, ignore_index=True)
            final_report.to_excel(dst, index=False)
            self._format_excel(dst)
            self._toggle_busy(False); self._set_status("Reporte generado exitosamente", "success")
            self._show_success_dialog(dst)
        except Exception as e:
            self._toggle_busy(False); self._set_status(f"Error: {e}", "error"); messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

    # ────────────  Formato Excel  ────────────
    def _format_excel(self, path):
        wb=load_workbook(path); ws=wb.active
        header_fill=PatternFill(start_color="3498DB", end_color="3498DB", fill_type="solid")
        header_font=Font(color="FFFFFF", bold=True)
        total_fill=PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        bold_font=Font(bold=True)
        thin=Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for c in range(1, ws.max_column+1):
            cell=ws.cell(1, c); cell.fill=header_fill; cell.font=header_font; cell.border=thin
        for r in range(2, ws.max_row+1):
            if ws.cell(r,2).value=='Totales':
                for c in range(1, ws.max_column+1):
                    cell=ws.cell(r,c); cell.fill=total_fill; cell.font=bold_font; cell.border=thin
        for col in ws.columns:
            length=max(len(str(cell.value)) if cell.value else 0 for cell in col)+2
            letter=col[0].column_letter
            if letter in ['A','B','C'] and length<15: length=15
            ws.column_dimensions[letter].width=length
        wb.save(path)

if __name__ == "__main__":
    root=Tk(); app=CheckadorApp(root); root.mainloop()
