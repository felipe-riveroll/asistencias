import os
import datetime
import subprocess
import traceback

from tkinter import (
    Tk,
    Label,
    Button,
    Entry,
    StringVar,
    filedialog,
    Frame,
    ttk,
    messagebox,
    Toplevel,
)
from PIL import Image, ImageTk

import pandas as pd

from expected_hours import load_expected_hours_data
from report import generate_report, format_excel

pd.options.mode.chained_assignment = None


class CheckadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Checadas")
        self.root.geometry("750x520")
        self.root.resizable(True, True)

        self.primary_color = "#2c3e50"
        self.secondary_color = "#3498db"
        self.bg_color = "#f5f5f5"
        self.text_color = "#333333"
        self.success_color = "#27ae60"
        self.warning_color = "#f39c12"
        self.error_color = "#e74c3c"
        self.orange_dark_color = "#E67E22"

        self.root.configure(bg=self.bg_color)

        self.expected_hours_df = load_expected_hours_data()
        self.expected_hours_cache: dict[str, float] | dict = {}

        self.status_frame = Frame(root, bg="#e0e0e0", relief="ridge", bd=1)
        self.status_frame.pack(side="bottom", fill="x")
        self.status_label = Label(
            self.status_frame,
            text="Listo para procesar",
            font=("Segoe UI", 10),
            bg="#e0e0e0",
            fg=self.primary_color,
            padx=10,
            pady=8,
        )
        self.status_label.pack(fill="x")

        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TEntry", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10), background=self.bg_color)

        self.input_file_path = StringVar()
        self.output_file_name = StringVar(
            value=f"reporte_checador_{datetime.datetime.now().strftime('%d%m%Y')}"
        )

        main = Frame(root, bg=self.bg_color, padx=30, pady=20)
        main.pack(fill="both", expand=True)

        try:
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logo_asia.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                w, h = img.size
                nw = 200
                nh = int(nw / w * h)
                logo = ImageTk.PhotoImage(img.resize((nw, nh), Image.LANCZOS))
                Label(main, image=logo, bg=self.bg_color).pack(pady=(0, 15))
                self._logo_ref = logo
        except Exception as e_logo:  # pragma: no cover - UI only
            print(f"Error loading logo: {e_logo}")

        banner = Frame(main, bg=self.primary_color, height=60)
        banner.pack(fill="x", pady=(0, 20))
        Label(
            banner,
            text="Procesador de Checadas de Empleados",
            font=("Segoe UI", 18, "bold"),
            bg=self.primary_color,
            fg="white",
        ).pack(pady=10)

        form = Frame(main, bg=self.bg_color, padx=20, pady=10)
        form.pack(fill="both", expand=True)
        row1 = Frame(form, bg=self.bg_color, pady=10)
        row1.pack(fill="x")
        Label(
            row1,
            text="Archivo Excel de Checadas:",
            font=("Segoe UI", 11),
            bg=self.bg_color,
            fg=self.text_color,
        ).pack(side="left", padx=(0, 10))
        Entry(row1, textvariable=self.input_file_path, font=("Segoe UI", 10), bd=1, relief="solid").pack(
            side="left", fill="x", expand=True, ipady=3
        )
        Button(
            row1,
            text="Examinar...",
            command=self.browse_file,
            font=("Segoe UI", 10),
            bg=self.secondary_color,
            fg="white",
            relief="flat",
        ).pack(side="left", padx=(10, 0))

        row2 = Frame(form, bg=self.bg_color, pady=10)
        row2.pack(fill="x")
        Label(
            row2,
            text="Nombre del archivo de salida:",
            font=("Segoe UI", 11),
            bg=self.bg_color,
            fg=self.text_color,
        ).pack(side="left", padx=(0, 10))
        Entry(row2, textvariable=self.output_file_name, font=("Segoe UI", 10), bd=1, relief="solid").pack(
            side="left", fill="x", expand=True, ipady=3
        )
        Label(row2, text=".xlsx", font=("Segoe UI", 11), bg=self.bg_color, fg=self.text_color).pack(
            side="left"
        )

        ttk.Separator(form, orient="horizontal").pack(fill="x", pady=20)

        actions = Frame(form, bg=self.bg_color, pady=20)
        actions.pack(fill="x")
        self.process_button = Button(
            actions,
            text="Procesar Archivo",
            command=self.generate_report,
            font=("Segoe UI", 12, "bold"),
            bg=self.secondary_color,
            fg="white",
            relief="flat",
            padx=20,
            pady=8,
        )
        self.process_button.pack(pady=10)
        self.progress = ttk.Progressbar(actions, orient="horizontal", length=500, mode="indeterminate")

    def browse_file(self):
        fp = filedialog.askopenfilename(
            filetypes=[("Archivos Excel", "*.xlsx;*.xls"), ("Todos", "*.*")]
        )
        if fp:
            self.input_file_path.set(fp)
            self.status_label.configure(
                text=f"Archivo seleccionado: {os.path.basename(fp)}",
                fg=self.primary_color,
                bg="#e0e0e0",
            )

    def _set_status(self, msg: str, kind: str = "info"):
        cfg = {
            "success": ("white", self.success_color),
            "warning": ("white", self.warning_color),
            "error": ("white", self.error_color),
            "info": (self.primary_color, "#e0e0e0"),
        }
        fg, bg = cfg.get(kind, cfg["info"])
        self.status_label.configure(text=msg, fg=fg, bg=bg)
        self.root.update()

    def _toggle_busy(self, busy: bool):
        if busy:
            self.process_button.configure(state="disabled", text="Procesando...", bg="#95a5a6")
            self.progress.pack(pady=10)
            self.progress.start(10)
        else:
            self.process_button.configure(state="normal", text="Procesar Archivo", bg=self.secondary_color)
            self.progress.stop()
            self.progress.pack_forget()

    def _show_success_dialog(self, path: str):
        dlg = Toplevel(self.root)
        dlg.title("Proceso completado")
        dlg.geometry("450x250")
        dlg.resizable(False, False)
        dlg.configure(bg="white")
        content = Frame(dlg, bg="white", padx=20, pady=10)
        content.pack(fill="both", expand=True)
        try:
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logo_asia.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                w, h = img.size
                nw = 150
                nh = int(nw / w * h)
                logo = ImageTk.PhotoImage(img.resize((nw, nh), Image.LANCZOS))
                Label(content, image=logo, bg="white").pack(pady=(0, 10))
                dlg._logo = logo
            else:
                Label(content, text="ⓘ", font=("Segoe UI", 24), fg=self.secondary_color, bg="white").pack(
                    pady=(10, 5)
                )
        except Exception:  # pragma: no cover - UI only
            Label(content, text="ⓘ", font=("Segoe UI", 24), fg=self.secondary_color, bg="white").pack(
                pady=(10, 5)
            )
        Label(
            content,
            text="El reporte se ha generado correctamente.",
            font=("Segoe UI", 11),
            bg="white",
        ).pack(pady=5)
        short = path if len(path) <= 50 else os.path.join(os.path.dirname(path)[:10] + "...", os.path.basename(path))
        pf = Frame(content, bg="white")
        pf.pack(pady=5, fill="x")
        Label(pf, text="Ubicación:", font=("Segoe UI", 10), bg="white").pack(side="left")
        Label(pf, text=short, font=("Segoe UI", 9), fg="#555", bg="white").pack(side="left")
        bf = Frame(content, bg="white")
        bf.pack(fill="x", pady=10)

        def _open_file_action():
            try:
                if os.name == "nt":
                    os.startfile(path)
                elif os.name == "posix":
                    subprocess.call(("open" if os.path.exists("/usr/bin/open") else "xdg-open", path))
                dlg.destroy()
            except Exception as e_open:  # pragma: no cover - UI only
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e_open}")

        Button(bf, text="Abrir archivo", font=("Segoe UI", 10), bg=self.secondary_color, fg="white", width=12, command=_open_file_action).pack(
            side="right", padx=5
        )
        Button(bf, text="Aceptar", font=("Segoe UI", 10), bg="#f0f0f0", width=10, command=dlg.destroy).pack(
            side="right", padx=5
        )
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.update_idletasks()
        w_dlg, h_dlg = dlg.winfo_width(), dlg.winfo_height()
        x = (dlg.winfo_screenwidth() - w_dlg) // 2
        y = (dlg.winfo_screenheight() - h_dlg) // 2
        dlg.geometry(f"{w_dlg}x{h_dlg}+{x}+{y}")

    def generate_report(self):
        src = self.input_file_path.get().strip()
        if not src:
            self._set_status("Seleccione un archivo", "error")
            messagebox.showerror("Error", "Debe seleccionar un archivo de entrada")
            return

        output_filename = self.output_file_name.get().strip() + ".xlsx"
        if os.path.isabs(src) and os.path.isdir(os.path.dirname(src)):
            dst_folder = os.path.dirname(src)
        else:
            dst_folder = os.path.dirname(os.path.abspath(__file__))
        dst = os.path.join(dst_folder, output_filename)

        try:
            self._toggle_busy(True)
            self._set_status("Procesando archivo...", "info")
            resumen_df = generate_report(src, dst, self.expected_hours_df, self.expected_hours_cache)
            format_excel(dst, resumen_df)
            self._toggle_busy(False)
            self._set_status("Reporte generado exitosamente", "success")
            self._show_success_dialog(dst)
        except ValueError as ve:
            self._toggle_busy(False)
            self._set_status(f"Error de valor: {ve}", "error")
            messagebox.showerror(
                "Error de Valor",
                f"Ocurrió un error con los datos:\n{ve}\n\n{traceback.format_exc()}",
            )
        except Exception as e:
            self._toggle_busy(False)
            self._set_status(f"Error: {e}", "error")
            messagebox.showerror(
                "Error",
                f"Ocurrió un error inesperado:\n{e}\n\n{traceback.format_exc()}",
            )

