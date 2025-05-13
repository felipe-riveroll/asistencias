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
    return hashlib.md5(pd.util.hash_pandas_object(dataframe).values).hexdigest()

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
    
    # Convertir en DataFrame y renombrar columnas
    df = pd.DataFrame(all_data)
    # Si tus columnas se llaman "# L", "# M", etc:
    rename_map = {
        "# L": "L", "# M": "M", "# X": "X",
        "# J": "J", "# V": "V", "# S": "S", "# D": "D"
    }
    df = df.rename(columns=rename_map)
    
    # Seleccionamos sólo Employee y días
    df = df[["Employee", "L","M","X","J","V","S","D"]]
    df["Employee"] = df["Employee"].astype(int)
    return df

def save_data_locally(df):
    """Guarda el DataFrame en un archivo local con metadatos"""
    # Guardar el DataFrame
    df.to_csv(LOCAL_DATA_FILE, index=False)
    
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
        return pd.read_csv(LOCAL_DATA_FILE)
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
    Devuelve un DataFrame con columnas: ['Employee','L','M','X','J','V','S','D']
    donde L=segundos esperados el lunes, M=martes, etc.
    """
    local_data = load_local_data()
    local_metadata = get_local_metadata()
    
    # Si no hay datos locales, descargar y guardar
    if local_data is None or local_metadata is None:
        print("No se encontró archivo local. Descargando datos...")
        df = fetch_data_from_api()
        save_data_locally(df)
        return df
    
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

# Ejemplo de uso:
if __name__ == "__main__":
    df_esperadas = fetch_expected_hours()
    # print(df_esperadas.head())
