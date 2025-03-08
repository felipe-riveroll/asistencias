import pandas as pd
import math
import os
import datetime
import subprocess
from tkinter import Tk, Label, Button, Entry, StringVar, filedialog, Frame, ttk, messagebox, Toplevel
from PIL import Image, ImageTk  # Para manejar imágenes
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Configurar pandas para mostrar las advertencias de asignación en cadena
pd.options.mode.chained_assignment = 'raise'

class CheckadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Checadas")
        self.root.geometry("750x520")  # Reducido al quitar el mensaje de descripción
        self.root.resizable(True, True)
        
        # Configurar colores y estilos
        self.primary_color = "#2c3e50"  # Azul oscuro
        self.secondary_color = "#3498db"  # Azul claro
        self.bg_color = "#f5f5f5"  # Gris muy claro
        self.text_color = "#333333"  # Gris oscuro
        self.success_color = "#27ae60"  # Verde
        self.warning_color = "#f39c12"  # Naranja
        self.error_color = "#e74c3c"  # Rojo
        
        self.root.configure(bg=self.bg_color)
        
        # Panel de estado - Creado PRIMERO para asegurar que esté en la parte inferior
        self.status_frame = Frame(root, bg="#e0e0e0", relief="ridge", bd=1)
        self.status_frame.pack(side="bottom", fill="x")
        
        self.status_label = Label(self.status_frame, text="Listo para procesar", 
                                font=("Segoe UI", 10), bg="#e0e0e0", fg=self.primary_color,
                                padx=10, pady=8)
        self.status_label.pack(fill="x")
        
        # Estilo para ttk
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TEntry", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10), background=self.bg_color)
        style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), foreground=self.primary_color, background=self.bg_color)
        style.configure("Subheader.TLabel", font=("Segoe UI", 12), foreground=self.text_color, background=self.bg_color)
        style.configure("Status.TLabel", font=("Segoe UI", 10, "bold"), background="#e0e0e0")
        
        # Variables para almacenar rutas
        self.input_file_path = StringVar()
        self.output_file_name = StringVar()
        self.processing = False
        
        # Contenedor principal
        main_container = Frame(root, bg=self.bg_color, padx=30, pady=20)
        main_container.pack(fill="both", expand=True)
        
        # Logo de la empresa
        try:
            # Cargar y redimensionar el logo
            logo_path = "Logo_asia.png"
            original_logo = Image.open(logo_path)
            
            # Calcular nueva altura manteniendo la proporción
            width, height = original_logo.size
            new_width = 200  # Ancho fijo deseado
            new_height = int((new_width / width) * height)
            
            resized_logo = original_logo.resize((new_width, new_height), Image.LANCZOS)
            logo_img = ImageTk.PhotoImage(resized_logo)
            
            # Crear y ubicar el logo
            logo_label = Label(main_container, image=logo_img, bg=self.bg_color)
            logo_label.image = logo_img  # Evitar que la imagen se elimine por el recolector de basura
            logo_label.pack(pady=(0, 15))
        except Exception as e:
            print(f"Error al cargar el logo: {e}")
        
        # Banner superior
        banner = Frame(main_container, bg=self.primary_color, height=60)
        banner.pack(fill="x", pady=(0, 20))
        
        title_label = Label(banner, text="Procesador de Checadas de Empleados", 
                           font=("Segoe UI", 18, "bold"), bg=self.primary_color, fg="white")
        title_label.pack(pady=10)
        
        # Contenedor de formulario
        form_container = Frame(main_container, bg=self.bg_color, padx=20, pady=10)
        form_container.pack(fill="both", expand=True)
        
        # Fila 1: Selección de archivo de entrada
        file_frame = Frame(form_container, bg=self.bg_color, pady=10)
        file_frame.pack(fill="x")
        
        file_label = Label(file_frame, text="Archivo Excel de Checadas:", 
                          font=("Segoe UI", 11), bg=self.bg_color, fg=self.text_color)
        file_label.pack(side="left", padx=(0, 10))
        
        file_entry = Entry(file_frame, textvariable=self.input_file_path, 
                          font=("Segoe UI", 10), bd=1, relief="solid")
        file_entry.pack(side="left", fill="x", expand=True, ipady=3)
        
        browse_button = Button(file_frame, text="Examinar...", command=self.browse_input_file,
                              font=("Segoe UI", 10), bg=self.secondary_color, fg="white",
                              relief="flat", padx=10, pady=2)
        browse_button.pack(side="left", padx=(10, 0))
        
        # Fila 2: Nombre de archivo de salida
        output_frame = Frame(form_container, bg=self.bg_color, pady=10)
        output_frame.pack(fill="x")
        
        output_label = Label(output_frame, text="Nombre del archivo de salida:", 
                            font=("Segoe UI", 11), bg=self.bg_color, fg=self.text_color)
        output_label.pack(side="left", padx=(0, 10))
        
        output_entry = Entry(output_frame, textvariable=self.output_file_name, 
                            font=("Segoe UI", 10), bd=1, relief="solid")
        output_entry.pack(side="left", fill="x", expand=True, ipady=3)
        
        extension_label = Label(output_frame, text=".xlsx", 
                              font=("Segoe UI", 11), bg=self.bg_color, fg=self.text_color)
        extension_label.pack(side="left")
        
        # Separador visual
        separator = ttk.Separator(form_container, orient="horizontal")
        separator.pack(fill="x", pady=20)
        
        # Sección de acciones
        action_frame = Frame(form_container, bg=self.bg_color, pady=20)
        action_frame.pack(fill="x")
        
        # Botón de procesamiento con efecto al presionar
        self.process_button = Button(action_frame, text="Procesar Archivo", command=self.generate_report,
                              font=("Segoe UI", 12, "bold"), bg=self.secondary_color, fg="white",
                              relief="flat", padx=20, pady=8, activebackground="#2980b9")
        self.process_button.pack(pady=10)
        
        # Barra de progreso
        self.progress_var = ttk.Progressbar(action_frame, orient="horizontal", 
                                          length=500, mode="indeterminate")
        
        # Línea separadora
        separator_bottom = ttk.Separator(form_container, orient="horizontal")
        separator_bottom.pack(fill="x", pady=(20, 10))
        
        # Establecer el nombre predeterminado para el archivo de salida
        today = datetime.datetime.now().strftime("%d%m%Y")
        self.output_file_name.set(f"reporte_checador_{today}")
    
    def browse_input_file(self):
        filetypes = [("Archivos Excel", "*.xlsx;*.xls"), ("Todos los archivos", "*.*")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        
        if file_path:
            self.input_file_path.set(file_path)
            
            # Actualizar automáticamente el nombre del archivo de salida en base a la fecha
            today = datetime.datetime.now().strftime("%d%m%Y")
            self.output_file_name.set(f"reporte_checador_{today}")
            
            self.status_label.config(text=f"Archivo seleccionado: {os.path.basename(file_path)}", 
                                    fg=self.primary_color, bg="#e0e0e0")
    
    def update_status(self, message, status_type="info"):
        """Actualizar la etiqueta de estado con el mensaje y color apropiado"""
        if status_type == "success":
            self.status_label.config(text=message, fg="white", bg=self.success_color)
        elif status_type == "warning":
            self.status_label.config(text=message, fg="white", bg=self.warning_color)
        elif status_type == "error":
            self.status_label.config(text=message, fg="white", bg=self.error_color)
        else:  # info
            self.status_label.config(text=message, fg=self.primary_color, bg="#e0e0e0")
        
        # Asegurar que la barra de estado sea visible
        self.status_frame.lift()
        self.status_label.lift()
        self.root.update()
    
    def toggle_processing_state(self, is_processing):
        """Alternar el estado de procesamiento de la interfaz"""
        if is_processing:
            self.process_button.config(state="disabled", text="Procesando...", bg="#95a5a6")
            self.progress_var.pack(pady=10)
            self.progress_var.start(10)
        else:
            self.process_button.config(state="normal", text="Procesar Archivo", bg=self.secondary_color)
            self.progress_var.stop()
            self.progress_var.pack_forget()
            
    def show_success_dialog(self, file_path):
        """Mostrar un diálogo personalizado con la opción de abrir el archivo"""
        # Crear un diálogo personalizado
        dialog = Toplevel(self.root)
        dialog.title("Proceso completado")
        dialog.geometry("450x250")  # Un poco más alto para el logo
        dialog.resizable(False, False)
        dialog.configure(bg="white")
        
        # Intentar establecer el icono de información
        try:
            dialog.iconbitmap("info.ico")  # Si tienes un ícono personalizado
        except:
            pass  # Si no hay ícono disponible, continuamos sin él
            
        # Crear un frame para el contenido
        content_frame = Frame(dialog, bg="white", padx=20, pady=10)
        content_frame.pack(fill="both", expand=True)
        
        # Intentar cargar el logo
        try:
            # Cargar y redimensionar el logo
            logo_path = "Logo_asia.png"
            original_logo = Image.open(logo_path)
            
            # Calcular nueva altura manteniendo la proporción
            width, height = original_logo.size
            new_width = 150  # Ancho fijo deseado
            new_height = int((new_width / width) * height)
            
            resized_logo = original_logo.resize((new_width, new_height), Image.LANCZOS)
            logo_img = ImageTk.PhotoImage(resized_logo)
            
            # Crear y ubicar el logo
            logo_label = Label(content_frame, image=logo_img, bg="white")
            logo_label.image = logo_img  # Evitar que la imagen se elimine por el recolector de basura
            logo_label.pack(pady=(0, 10))
        except Exception as e:
            # Si no se puede cargar el logo, mostramos el ícono de información
            info_label = Label(content_frame, text="ⓘ", font=("Segoe UI", 24), fg="#3498db", bg="white")
            info_label.pack(pady=(10, 5))
        
        # Mensaje de éxito
        message_label = Label(content_frame, text="El reporte se ha generado correctamente.", 
                             font=("Segoe UI", 11), bg="white")
        message_label.pack(pady=5)
        
        # Mostrar la ubicación del archivo
        path_frame = Frame(content_frame, bg="white")
        path_frame.pack(fill="x", pady=5)
        
        path_label = Label(path_frame, text="Ubicación:", font=("Segoe UI", 10), bg="white")
        path_label.pack(side="left", padx=(0, 5))
        
        # Truncar la ruta si es muy larga
        path_display = file_path
        if len(path_display) > 50:
            parts = path_display.split('/')
            if len(parts) > 1:
                path_display = '/'.join(parts[:1] + ['...'] + parts[-1:])
        
        file_path_label = Label(path_frame, text=path_display, font=("Segoe UI", 9), 
                              fg="#555555", bg="white")
        file_path_label.pack(side="left")
        
        # Frame para botones
        buttons_frame = Frame(content_frame, bg="white")
        buttons_frame.pack(fill="x", pady=10)
        
        # Función para abrir el archivo Excel
        def open_excel_file():
            try:
                import subprocess
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                elif os.name == 'posix':  # macOS y Linux
                    if os.path.exists('/usr/bin/open'):  # macOS
                        subprocess.call(('open', file_path))
                    else:  # Linux
                        subprocess.call(('xdg-open', file_path))
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo: {str(e)}")
        
        # Botón para cerrar el diálogo
        accept_button = Button(buttons_frame, text="Aceptar", font=("Segoe UI", 10), 
                             bg="#f0f0f0", width=10, command=dialog.destroy)
        accept_button.pack(side="right", padx=5)
        
        # Botón para abrir el archivo
        open_button = Button(buttons_frame, text="Abrir archivo", font=("Segoe UI", 10), 
                           bg="#3498db", fg="white", width=12, command=open_excel_file)
        open_button.pack(side="right", padx=5)
        
        # Centrar el diálogo en la pantalla
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        # Hacer que el diálogo sea modal
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Asegurar que el diálogo se cierre correctamente
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
    
    def generate_report(self):
        input_path = self.input_file_path.get()
        output_name = self.output_file_name.get() + ".xlsx"
        
        if not input_path:
            self.update_status("Error: Por favor seleccione un archivo de entrada", "error")
            messagebox.showerror("Error", "Por favor seleccione un archivo de entrada")
            return
        
        # Ubicar el archivo de salida en la misma carpeta que el archivo de entrada
        output_directory = os.path.dirname(input_path)
        output_path = os.path.join(output_directory, output_name)
        
        try:
            # Activar estado de procesamiento
            self.toggle_processing_state(True)
            self.update_status("Procesando archivo...", "info")
            
            # Función para formatear un timedelta a "HH:MM:SS"
            def format_timedelta(td):
                # Convertir el timedelta a segundos totales, redondeando a entero
                total_seconds = int(round(td.total_seconds()))
                hours = total_seconds // 3600
                minutes = (total_seconds % 3600) // 60
                seconds = total_seconds % 60
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

            # Cargar el archivo Excel
            df = pd.read_excel(input_path)
            
            # Verificar las columnas disponibles en el archivo
            column_list = df.columns.tolist()
            self.update_status(f"Analizando estructura del archivo...", "info")
            
            # Verificar si existe la columna 'Time'
            if 'Time' not in df.columns:
                # Buscar columnas que puedan contener información de tiempo/fecha
                time_column = None
                for col in df.columns:
                    # Intentar encontrar una columna de tiempo basada en el nombre
                    if any(time_word in col.lower() for time_word in ['time', 'fecha', 'hora', 'date']):
                        time_column = col
                        break
                
                if time_column is None:
                    # Si no se encuentra una columna de tiempo por nombre, intentar detectar por tipo de datos
                    for col in df.columns:
                        # Verificar si la columna parece contener fechas
                        try:
                            if pd.to_datetime(df[col], errors='coerce').notna().any():
                                time_column = col
                                break
                        except:
                            continue
                
                if time_column is None:
                    raise ValueError("No se encontró una columna de tiempo en el archivo. Asegúrate de que el archivo contenga una columna con fechas y horas.")
                
                # Informar al usuario qué columna se está usando
                self.update_status(f"Usando '{time_column}' como columna de tiempo", "info")
                
                # Renombrar la columna para que el resto del código funcione
                df = df.rename(columns={time_column: 'Time'})
            
            # Convertir la columna 'Time' a tipo datetime con manejo de errores
            try:
                df = df.assign(Time=pd.to_datetime(df['Time'], errors='coerce'))
                
                # Verificar si hay valores NaN después de la conversión
                if df['Time'].isna().any():
                    num_nan = df['Time'].isna().sum()
                    self.update_status(f"Advertencia: {num_nan} registros con formato de fecha/hora inválido", "warning")
                    
                    # Eliminar filas con valores NaN en 'Time'
                    df = df.dropna(subset=['Time'])
                    if df.empty:
                        raise ValueError("Todos los registros tenían formatos de fecha/hora inválidos.")
            except Exception as e:
                raise ValueError(f"Error al convertir la columna 'Time' a datetime: {str(e)}")
            
            # Crear columna 'Day' a partir de la fecha
            df = df.assign(Day=df['Time'].dt.date)
            
            # Verificar si existe la columna 'Shift'
            if 'Shift' not in df.columns:
                # Buscar columnas que puedan contener información de turno
                shift_column = None
                for col in df.columns:
                    # Intentar encontrar una columna de turno basada en el nombre
                    if any(shift_word in col.lower() for shift_word in ['shift', 'turno', 'jornada']):
                        shift_column = col
                        break
                
                if shift_column is None:
                    # Si no podemos encontrar una columna adecuada, creamos una vacía
                    self.update_status("No se encontró columna de turno. Se creará una columna vacía.", "warning")
                    df['Shift'] = ""
                else:
                    # Renombrar la columna para que el resto del código funcione
                    df = df.rename(columns={shift_column: 'Shift'})
                    self.update_status(f"Usando '{shift_column}' como columna de turno", "info")
            
            # Verificar si existe la columna 'Employee Name'
            if 'Employee Name' not in df.columns:
                # Buscar columnas que puedan contener información de empleados
                employee_column = None
                for col in df.columns:
                    # Intentar encontrar una columna de empleados basada en el nombre
                    if any(emp_word in col.lower() for emp_word in ['employee', 'empleado', 'name', 'nombre']):
                        employee_column = col
                        break
                
                if employee_column is None:
                    raise ValueError("No se encontró una columna de empleados en el archivo. Asegúrate de que el archivo contenga una columna con nombres de empleados.")
                
                # Renombrar la columna para que el resto del código funcione
                df = df.rename(columns={employee_column: 'Employee Name'})
                self.update_status(f"Usando '{employee_column}' como columna de empleados", "info")
            
            self.update_status("Procesando registros de checadas...", "info")
            
            # Crear una copia del DataFrame original para conservar la información de los registros sin turno
            df_original = df.copy()
            
            # Primero, procesamos los registros que SÍ tienen turno asignado
            df_with_shift = df[df['Shift'].notna() & (df['Shift'] != "")].copy()
            
            # Identificamos los registros sin turno para procesarlos más tarde
            df_without_shift = df[df['Shift'].isna() | (df['Shift'] == "")].copy()
            df_without_shift['Merged'] = False  # Bandera para rastrear si el registro se ha fusionado
            
            # Manejar valores NaN o vacíos en 'Shift' para el DataFrame que los contiene
            df_with_shift['Shift'] = df_with_shift['Shift'].fillna("").astype(str)
            
            # Ordenar los datos por 'Employee Name', 'Shift', 'Day' y 'Time'
            df_with_shift = df_with_shift.sort_values(['Employee Name', 'Shift', 'Day', 'Time'])
            
            # Agrupar por 'Employee Name', 'Shift' y 'Day' para los registros con turno
            grouped_with_shift = df_with_shift.groupby(['Employee Name', 'Shift', 'Day']).agg(
                checadas_list=('Time', lambda times: sorted(list(times))),
                total_timedelta=('Time', lambda times: max(times) - min(times))
            ).reset_index()
            
            # Guardar registro de las celdas que deben marcarse en amarillo (tiempo, empleado, día, índice)
            cells_to_highlight = []
            
            # Función para procesar y fusionar los registros sin turno
            def merge_no_shift_records(grouped_data, no_shift_data):
                # Crear una copia para no modificar el original mientras iteramos
                result_data = grouped_data.copy()
                
                # Para cada registro sin turno
                for idx, row in no_shift_data.iterrows():
                    employee_name = row['Employee Name']
                    check_day = row['Day']
                    check_time = row['Time']
                    
                    # Buscar si hay registros para este empleado en este día
                    matching_rows = result_data[
                        (result_data['Employee Name'] == employee_name) & 
                        (result_data['Day'] == check_day)
                    ]
                    
                    if len(matching_rows) > 0:
                        # Hay un registro existente para este empleado en este día
                        for match_idx, match_row in matching_rows.iterrows():
                            # Obtener y actualizar la lista de checadas
                            current_checadas = match_row['checadas_list']
                            if check_time not in current_checadas:
                                # Marcar esta celda para resaltarla en amarillo
                                # Encontrar el índice que tendrá esta checada en la lista ordenada
                                new_checadas = sorted(current_checadas + [check_time])
                                checada_idx = new_checadas.index(check_time) + 1  # +1 porque las columnas empiezan en "Checada 1"
                                
                                cells_to_highlight.append({
                                    'employee': employee_name,
                                    'day': check_day,
                                    'shift': match_row['Shift'],
                                    'column': f'Checada {checada_idx}'
                                })
                                
                                # Actualizar el registro
                                result_data.at[match_idx, 'checadas_list'] = new_checadas
                                
                                # Actualizar el timedelta total y efectivo
                                if min(new_checadas) < min(current_checadas) or max(new_checadas) > max(current_checadas):
                                    result_data.at[match_idx, 'total_timedelta'] = max(new_checadas) - min(new_checadas)
                                
                                # Marcar como fusionado
                                no_shift_data.at[idx, 'Merged'] = True
                    
                # Identificar los registros sin turno que no se fusionaron
                unmergeable = no_shift_data[~no_shift_data['Merged']]
                
                # Proceso para registros que no pudieron ser fusionados
                if len(unmergeable) > 0:
                    # Procesamos los registros sin turno que no pudieron ser fusionados
                    unmergeable_grouped = unmergeable.groupby(['Employee Name', 'Day']).agg(
                        checadas_list=('Time', lambda times: sorted(list(times))),
                        total_timedelta=('Time', lambda times: max(times) - min(times))
                    ).reset_index()
                    
                    # Añadir columna 'Shift' vacía
                    unmergeable_grouped['Shift'] = ""
                    
                    # Reordenar columnas para que coincidan
                    unmergeable_grouped = unmergeable_grouped[['Employee Name', 'Shift', 'Day', 'checadas_list', 'total_timedelta']]
                    
                    # Concatenar con el resultado
                    result_data = pd.concat([result_data, unmergeable_grouped], ignore_index=True)
                
                return result_data
            
            # Fusionar los registros sin turno
            merged_grouped = merge_no_shift_records(grouped_with_shift, df_without_shift)
            
            # Ordenar el resultado
            merged_grouped = merged_grouped.sort_values(['Employee Name', 'Day'])
            
            # Función para calcular el timedelta de descanso
            def compute_break_timedelta(checadas):
                if len(checadas) == 4:
                    # Se asume que el segundo registro es el inicio y el tercero el fin del descanso
                    return checadas[2] - checadas[1]
                else:
                    return pd.Timedelta(0)

            # Calcular el timedelta de descanso y el de horas efectivas usando assign
            merged_grouped = merged_grouped.assign(
                break_timedelta=merged_grouped['checadas_list'].apply(compute_break_timedelta),
            )
            
            # Calcular el effective_timedelta
            merged_grouped = merged_grouped.assign(
                effective_timedelta=merged_grouped['total_timedelta'] - merged_grouped['break_timedelta']
            )

            # Formatear los timdeltas a cadenas en formato "HH:MM:SS"
            merged_grouped = merged_grouped.assign(
                Horas_totales=merged_grouped['total_timedelta'].apply(format_timedelta),
                Horas_de_descanso=merged_grouped['break_timedelta'].apply(format_timedelta),
                Horas_efectivas=merged_grouped['effective_timedelta'].apply(format_timedelta),
                Checadas_str=merged_grouped['checadas_list'].apply(lambda times: [t.strftime('%H:%M:%S') for t in times])
            )
            
            # Determinar el número máximo de checadas
            max_checadas = max(merged_grouped['Checadas_str'].apply(len))
            
            # Expandir la lista de checadas en columnas separadas
            checadas_columns = {}
            for i in range(max_checadas):
                col_name = f'Checada {i+1}'
                checadas_columns[col_name] = merged_grouped['Checadas_str'].apply(
                    lambda x: x[i] if i < len(x) else None
                )
            
            checadas_df = pd.DataFrame(checadas_columns)

            # Combinar la información en un solo DataFrame
            report = pd.concat([
                merged_grouped[['Employee Name', 'Shift', 'Day', 'Horas_totales', 'Horas_de_descanso', 'Horas_efectivas']], 
                checadas_df
            ], axis=1)

            # Renombrar columnas para el reporte final
            report = report.rename(columns={
                'Employee Name': 'Nombre del empleado',
                'Shift': 'Turno',
                'Day': 'Día',
                'Horas_totales': 'Horas totales',
                'Horas_de_descanso': 'Horas de descanso',
                'Horas_efectivas': 'Horas efectivas'
            })

            # --- Generar filas de totales por empleado ---
            # Primero, agrupar sobre el DataFrame merged_grouped (con time deltas sin formatear)
            totals = merged_grouped.groupby('Employee Name').agg({
                'total_timedelta': 'sum',
                'break_timedelta': 'sum',
                'effective_timedelta': 'sum'
            }).reset_index()

            # Formatear los totales a cadena
            totals = totals.assign(
                Horas_totales=totals['total_timedelta'].apply(format_timedelta),
                Horas_de_descanso=totals['break_timedelta'].apply(format_timedelta),
                Horas_efectivas=totals['effective_timedelta'].apply(format_timedelta)
            )

            # Crear filas de totales con las mismas columnas que report
            total_rows = []
            # Se recorren cada uno de los empleados
            for _, row in totals.iterrows():
                total_row = {
                    'Nombre del empleado': row['Employee Name'],
                    'Turno': 'Totales',
                    'Día': '',
                    'Horas totales': row['Horas_totales'],
                    'Horas de descanso': row['Horas_de_descanso'],
                    'Horas efectivas': row['Horas_efectivas']
                }
                # Para las columnas de Checada, dejar vacío
                for col in report.columns:
                    if col.startswith("Checada"):
                        total_row[col] = ""
                total_rows.append(total_row)

            totals_df = pd.DataFrame(total_rows)

            # Insertar la fila de totales al final de cada grupo de empleado
            final_rows = []
            # Agrupar report por "Nombre del empleado" en el orden original
            for emp, group_df in report.groupby('Nombre del empleado', sort=False):
                final_rows.append(group_df)
                # Obtener la fila de totales para este empleado
                total_row = totals_df[totals_df['Nombre del empleado'] == emp]
                final_rows.append(total_row)
                
            final_report = pd.concat(final_rows, ignore_index=True)

            # Guardar un registro de qué filas son totales para formatearlas después
            total_rows_indices = []
            current_row = 0
            for emp, group_df in report.groupby('Nombre del empleado', sort=False):
                current_row += len(group_df)
                total_rows_indices.append(current_row)
                current_row += 1  # La fila de totales

            self.update_status("Aplicando formato al reporte...", "info")
            
            # Exportar el reporte final a un nuevo archivo Excel
            final_report.to_excel(output_path, index=False)

            # Ahora, aplicar formato a las filas de totales y ajustar el ancho de las columnas
            wb = load_workbook(output_path)
            ws = wb.active

            # Definir el formato para las celdas de totales
            bold_font = Font(bold=True)
            fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Color gris claro
            thin_border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
            # Definir el color amarillo para las celdas fusionadas
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

            # Aplicar formato a las filas de totales
            for idx in total_rows_indices:
                # La fila en Excel será idx + 2 (1 por índice base 0 de pandas y 1 por cabeceras)
                excel_row = idx + 2
                
                # Aplicar formato a todas las celdas de la fila
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=excel_row, column=col)
                    cell.font = bold_font
                    cell.fill = fill
                    cell.border = thin_border
            
            # Aplicar resaltado amarillo a las celdas de los registros fusionados
            # Primero convertir el DataFrame final_report a un diccionario que permita búsquedas rápidas
            report_dict = {}
            for idx, row in final_report.iterrows():
                employee = row['Nombre del empleado']
                shift = row['Turno']
                day = row['Día']
                report_dict[(employee, day, shift)] = idx
            
            # Resaltar las celdas marcadas para ser destacadas
            for highlight in cells_to_highlight:
                employee = highlight['employee']
                day = highlight['day']
                shift = highlight['shift']
                column = highlight['column']
                
                # Buscar la fila en el reporte
                key = (employee, day, shift)
                if key in report_dict:
                    row_idx = report_dict[key] + 2  # +2 por la cabecera y el índice 0-base
                    col_idx = list(final_report.columns).index(column) + 1  # +1 porque las columnas en openpyxl son 1-base
                    
                    # Aplicar formato amarillo
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.fill = yellow_fill

            # Ajustar el ancho de las columnas según su contenido
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter  # Obtener la letra de la columna
                
                # Encontrar la longitud máxima en todas las celdas de esta columna
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                # Añadir un margen para mejor visualización
                adjusted_width = max_length + 2
                
                # Establecer el ancho de la columna
                ws.column_dimensions[column].width = adjusted_width
                
                # Asegurar un ancho mínimo para las primeras tres columnas
                if column in ['A', 'B', 'C']:
                    min_width = 15  # Definir un ancho mínimo para estas columnas
                    if adjusted_width < min_width:
                        ws.column_dimensions[column].width = min_width

            # Agregar cabeceros con formato
            header_fill = PatternFill(start_color="3498DB", end_color="3498DB", fill_type="solid")  # Azul
            header_font = Font(bold=True, color="FFFFFF")  # Texto blanco en negrita
            
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin_border

            # Guardar el archivo con el formato aplicado
            wb.save(output_path)
            
            # Desactivar estado de procesamiento
            self.toggle_processing_state(False)
            
            # Actualizar la etiqueta de estado
            self.update_status("Reporte generado exitosamente", "success")
            
            # Mostrar un diálogo personalizado con botón para abrir el archivo
            self.show_success_dialog(output_path)
            
        except Exception as e:
            # Desactivar estado de procesamiento
            self.toggle_processing_state(False)
            
            error_message = str(e)
            self.update_status(f"Error: {error_message}", "error")
            messagebox.showerror("Error", f"Ha ocurrido un error durante el procesamiento:\n\n{error_message}")

if __name__ == "__main__":
    root = Tk()
    app = CheckadorApp(root)
    root.mainloop()