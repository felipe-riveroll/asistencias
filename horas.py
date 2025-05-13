import os
from dotenv import load_dotenv
import requests
import pandas as pd
import hashlib
import json
from datetime import datetime

# Carga variables de entorno
load_dotenv()  # Carga las variables desde .env
API_URL     = os.getenv("NOCODB_API_URL")
API_KEY     = os.getenv("NOCODB_API_KEY")
PROJECT     = os.getenv("NOCODB_PROJECT_ID")
TABLE       = os.getenv("NOCODB_TABLE_NAME")

# Nombre del archivo local para almacenar los datos
LOCAL_DATA_FILE = "expected_hours_data.csv"
LOCAL_METADATA_FILE = "expected_hours_metadata.json"

def get_data_hash(dataframe):
    """Genera un hash único para el contenido del DataFrame"""
    # Hacemos una copia para no modificar el DataFrame original
    df_copy = dataframe.copy()
    
    # Normalizar nombres de columnas para comparación consistente
    # (convertir a formato corto para comparación)
    day_names_map = {
        "Lunes": "L",
        "Martes": "M",
        "Miércoles": "X",
        "Jueves": "J",
        "Viernes": "V",
        "Sábado": "S",
        "Domingo": "D"
    }
    
    # Renombrar solo las columnas que existen y coinciden con el mapeo
    rename_cols = {col: day_names_map.get(col, col) for col in df_copy.columns if col in day_names_map}
    df_copy = df_copy.rename(columns=rename_cols)
    
    # Generar el hash con los nombres normalizados
    return hashlib.md5(pd.util.hash_pandas_object(df_copy).values).hexdigest()

def fetch_data_from_api():
    """Descarga datos desde la API y los transforma en DataFrame usando paginación"""
    headers = {
        "xc-token": API_KEY,
        "Content-Type": "application/json"
    }
    url = f"{API_URL}/api/v2/tables/{TABLE}/records"
    
    all_data = []
    page_size = 100
    offset = 0
    
    # Bucle para obtener todos los registros usando paginación
    while True:
        params = {
            "offset": offset,
            "limit": page_size,
            "where": "",
            "viewId": PROJECT
        }
        
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        page_data = resp.json()["list"]
        
        # Añadir datos a la lista acumulativa
        all_data.extend(page_data)
        
        # Si recibimos menos registros que el tamaño de página, hemos terminado
        if len(page_data) < page_size:
            break
              # Actualizar el offset para la siguiente página
        offset += page_size
    
    print(f"Total de registros descargados: {len(all_data)}")
    
    # Convertir en DataFrame
    df = pd.DataFrame(all_data)
    
    # Imprimir nombres de columnas para diagnóstico
    print("Columnas en el DataFrame:", df.columns.tolist())
    
    # Verificar si las columnas esperadas están presentes
    # O si están con un formato diferente
    day_columns = ["# L", "# M", "# X", "# J", "# V", "# S", "# D"]
    
    # Mapeo para diferentes posibles nombres de columnas
    possible_columns = {
        "# L": ["# L", "#L", "L", "Lunes"],
        "# M": ["# M", "#M", "M", "Martes"],
        "# X": ["# X", "#X", "X", "Miércoles"],
        "# J": ["# J", "#J", "J", "Jueves"],
        "# V": ["# V", "#V", "V", "Viernes"],
        "# S": ["# S", "#S", "S", "Sábado"],
        "# D": ["# D", "#D", "D", "Domingo"],
    }
      # Crear el mapeo dinámico de columnas (formato corto para procesamiento interno)
    rename_map = {}
    for day, variants in possible_columns.items():
        for variant in variants:
            if variant in df.columns:
                rename_map[variant] = day.replace("# ", "")
                break
    
    print("Mapeo de columnas:", rename_map)
    
    # Renombrar columnas (formato corto para procesamiento interno)
    df = df.rename(columns=rename_map)
    
    # Asegurarse de que existe una columna Employee
    if "Employee" not in df.columns:
        # Buscar columnas que puedan contener IDs de empleados
        employee_column = None
        for col in df.columns:
            if "employee" in col.lower() or "id" in col.lower():
                employee_column = col
                break
        
        if employee_column:
            df = df.rename(columns={employee_column: "Employee"})
        else:
            raise ValueError("No se encontró una columna que identifique a los empleados")
    
    # Seleccionar las columnas necesarias
    required_columns = ["Employee"]
    available_day_columns = []
    
    for day in ["L", "M", "X", "J", "V", "S", "D"]:
        if day in df.columns:
            available_day_columns.append(day)
    
    # Verificar si tenemos al menos algunas columnas de días
    if not available_day_columns:
        print("¡ADVERTENCIA! No se encontraron columnas de días. Usando columnas disponibles.")
        print("Columnas disponibles:", df.columns.tolist())
        # Usar algunas columnas numéricas como días si están disponibles
        for col in df.columns:
            if col != "Employee" and col not in required_columns:
                try:
                    # Intentar convertir a número para verificar si es una columna numérica
                    pd.to_numeric(df[col], errors='raise')
                    available_day_columns.append(col)
                    if len(available_day_columns) >= 7:  # Máximo 7 días
                        break
                except:
                    pass
    
    # Seleccionar columnas
    columns_to_select = required_columns + available_day_columns
    print("Columnas seleccionadas:", columns_to_select)
    
    df = df[columns_to_select]
    
    # Convertir Employee a entero si es posible
    try:
        df["Employee"] = df["Employee"].astype(int)
    except:
        print("No se pudo convertir la columna Employee a entero")
    return df

def save_data_locally(df):
    """Guarda el DataFrame en un archivo local con metadatos"""
    # Crear una copia del DataFrame para no modificar el original
    df_to_save = df.copy()
    
    # Renombrar las columnas con nombres de días completos antes de guardar
    day_names_map = {
        "L": "Lunes",
        "M": "Martes",
        "X": "Miércoles",
        "J": "Jueves",
        "V": "Viernes",
        "S": "Sábado",
        "D": "Domingo"
    }
    
    # Aplicar el renombramiento solo a las columnas que existen
    rename_cols = {col: day_names_map.get(col, col) for col in df_to_save.columns if col in day_names_map}
    df_to_save = df_to_save.rename(columns=rename_cols)
    
    # Guardar el DataFrame con nombres de columnas completos
    df_to_save.to_csv(LOCAL_DATA_FILE, index=False)
    print(f"Columnas guardadas en el archivo: {df_to_save.columns.tolist()}")
    
    # Guardar metadatos (fecha de actualización y hash)
    metadata = {
        "last_update": datetime.now().isoformat(),
        "data_hash": get_data_hash(df)
    }
    
    with open(LOCAL_METADATA_FILE, 'w') as f:
        json.dump(metadata, f)
    
    print(f"Datos guardados localmente en {LOCAL_DATA_FILE}")

def load_local_data():
    """Carga datos desde el archivo local si existe"""
    if os.path.exists(LOCAL_DATA_FILE):
        df = pd.read_csv(LOCAL_DATA_FILE)
        # El archivo CSV usa nombres completos, pero internamente utilizamos abreviaturas
        # Para mantener consistencia, los convertimos de nuevo a formato corto
        day_names_map = {
            "Lunes": "L",
            "Martes": "M",
            "Miércoles": "X",
            "Jueves": "J",
            "Viernes": "V",
            "Sábado": "S",
            "Domingo": "D"
        }
        
        # Cuando cargamos el archivo local, dejamos los nombres como están 
        # (completos) para mantener consistencia con el archivo guardado
        return df
    return None

def get_local_metadata():
    """Obtiene los metadatos del archivo local si existe"""
    if os.path.exists(LOCAL_METADATA_FILE):
        with open(LOCAL_METADATA_FILE, 'r') as f:
            return json.load(f)
    return None

def fetch_expected_hours():
    """
    Descarga registros de la tabla o usa la versión local si está disponible y actualizada.
    Devuelve un DataFrame con la columna 'Employee' y las columnas de días disponibles
    donde cada columna contiene los segundos esperados para ese día.
    """
    local_data = load_local_data()
    local_metadata = get_local_metadata()
    
    # Si no hay datos locales, descargar y guardar
    if local_data is None or local_metadata is None:
        print("No se encontró archivo local. Descargando datos...")
        df = fetch_data_from_api()
        save_data_locally(df)
        return df
    
    try:
        # Comprobar si hay cambios comparando con la API
        print("Verificando si hay actualizaciones...")
        api_data = fetch_data_from_api()
        api_hash = get_data_hash(api_data)
        
        if api_hash != local_metadata.get('data_hash'):
            print("Se detectaron cambios en los datos. Actualizando archivo local...")
            save_data_locally(api_data)
            return api_data
        else:
            print(f"Usando datos locales (última actualización: {local_metadata.get('last_update')})")
            return local_data
    except Exception as e:
        print(f"Error al comprobar actualizaciones: {str(e)}")
        print("Usando datos locales como respaldo...")
        if local_data is not None:
            return local_data
        else:
            raise Exception("No se pudieron obtener datos de la API ni localmente")

# Ejemplo de uso:
if __name__ == "__main__":
    df_esperadas = fetch_expected_hours()
    print("\n--- Datos obtenidos ---")
    print(f"Total de filas: {len(df_esperadas)}")
    print(f"Columnas disponibles: {df_esperadas.columns.tolist()}")
    
    # Mostrar los datos con nombres completos para días de la semana
    df_display = df_esperadas.copy()
    day_names_map = {
        "L": "Lunes",
        "M": "Martes",
        "X": "Miércoles",
        "J": "Jueves",
        "V": "Viernes",
        "S": "Sábado",
        "D": "Domingo"
    }
    rename_cols = {col: day_names_map.get(col, col) for col in df_display.columns if col in day_names_map}
    
    if rename_cols:  # Solo renombrar si hay columnas para cambiar
        df_display = df_display.rename(columns=rename_cols)
        print("\nDatos con nombres completos de días:")
    
    print("\nPrimeras 5 filas:")
    print(df_display.head())
