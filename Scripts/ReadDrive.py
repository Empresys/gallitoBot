import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
import io
from googleapiclient.http import MediaIoBaseDownload

# Ruta del archivo Excel local que usas como referencia
ruta_archivo_excel = '../TEMPORAL/datos_factura.xlsx'

# Leer el archivo de Excel
try:
    df = pd.read_excel(ruta_archivo_excel)
    print("Archivo de Excel leído correctamente.")
except Exception as e:
    print(f"Error al leer el archivo de Excel: {e}")
    exit()

# Verificar que las columnas 'Sucursal' y 'Fecha' existan
if 'Sucursal' not in df.columns or 'Fecha' not in df.columns:
    print("Las columnas 'Sucursal' o 'Fecha' no se encuentran en el archivo de Excel.")
    exit()

# Convertir la columna 'Fecha' al formato deseado (por ejemplo, "YYYY-MM-DD")
df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%d')

# Obtener la lista de sucursales y fechas con el nuevo formato
sucursal_columna = df['Sucursal'].tolist()
fecha_columna = df['Fecha'].tolist()

# Cargar credenciales desde el archivo JSON
SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'credentials.json'  # Cambia esto a la ruta de tu archivo JSON

# Crear las credenciales
try:
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    print("Credenciales cargadas correctamente.")
except Exception as e:
    print(f"Error al cargar las credenciales: {e}")
    exit()

# Construir el servicio
service = build('drive', 'v3', credentials=credentials)

# IDs de las carpetas de sucursales
carpetas_sucursales = {
    'CANTERA': '1IJ5pOa5sNxeBDAPtjc3Z3NzH6JKHE9RW',
    'CUAUHTEMOC': '1fgp5t-4GQRII2zjas6h2KkUhp93iRyau',
    'NORTE': '1HDWEcVWzQW6aemDM2Cv5WGBCSIXol_ic'
}

# Función para buscar y descargar un archivo de Google Drive basado en el nombre (considerando las extensiones)
# Función para buscar y descargar un archivo de Google Drive basado en el nombre exacto
def descargar_archivo(nombre_archivo_base, carpeta_id):
    # Normalizar el nombre del archivo eliminando espacios adicionales y convirtiendo a minúsculas
    nombre_archivo_base = nombre_archivo_base.strip().lower()

    # Crear las consultas para ambos formatos
    query_xls = f"'{carpeta_id}' in parents and name = '{nombre_archivo_base}.xls'"
    query_xlsx = f"'{carpeta_id}' in parents and name = '{nombre_archivo_base}.xlsx'"

    # Intentar buscar el archivo con nombre completo incluyendo la extensión .xls
    results_xls = service.files().list(
        q=query_xls,
        pageSize=1,  # Buscar solo el primer resultado exacto
        fields="files(id, name)"
    ).execute()
    
    items = results_xls.get('files', [])

    # Si no se encontró con .xls, intentar con .xlsx
    if not items:
        results_xlsx = service.files().list(
            q=query_xlsx,
            pageSize=1,  # Buscar solo el primer resultado exacto
            fields="files(id, name)"
        ).execute()
        items = results_xlsx.get('files', [])
    
    if not items:
        print(f'No se encontró el archivo exacto "{nombre_archivo_base}" con las extensiones .xls o .xlsx en la carpeta especificada.')
        return
    else:
        # Descargar el archivo encontrado
        archivo_seleccionado = items[0]
        file_id = archivo_seleccionado['id']
        nombre_archivo_encontrado = archivo_seleccionado['name']
        
        request = service.files().get_media(fileId=file_id)
        destino_local = f"../TEMPORAL/{nombre_archivo_encontrado}"  # Actualizar la ruta según el nombre encontrado
        fh = io.FileIO(destino_local, 'wb')
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = downloader.next_chunk()
            print(f"Descargando {nombre_archivo_encontrado}: {int(status.progress() * 100)}%")

        print(f"Archivo {nombre_archivo_encontrado} descargado correctamente en {destino_local}.")

# Descargar archivos según las sucursales y fechas encontradas
for i in range(len(sucursal_columna)):
    sucursal = sucursal_columna[i]
    fecha = fecha_columna[i]

    # Verificar si la sucursal tiene una carpeta asignada
    if sucursal in carpetas_sucursales:
        carpeta_id = carpetas_sucursales[sucursal]
        nombre_archivo_base = f"{sucursal} {fecha}"  # Formato base para buscar, ignorando la extensión

        print(f"Descargando archivo que coincida exactamente con '{nombre_archivo_base}' desde la carpeta {sucursal} (ID: {carpeta_id})")

        # Descargar archivo
        descargar_archivo(nombre_archivo_base, carpeta_id)
    else:
        print(f"Sucursal {sucursal} no tiene carpeta asignada.")

    # Normalizar el nombre del archivo eliminando espacios adicionales y convirtiendo a minúsculas
    nombre_archivo_base = nombre_archivo_base.strip().lower()

    # Crear las consultas para ambos formatos
    query_sin_extension = f"'{carpeta_id}' in parents and name contains '{nombre_archivo_base}'"
    query_xls = f"'{carpeta_id}' in parents and name = '{nombre_archivo_base}.xls'"
    query_xlsx = f"'{carpeta_id}' in parents and name = '{nombre_archivo_base}.xlsx'"

    # Buscar archivos que contengan el nombre base (sucursal y fecha), ignorando la extensión
    print(f"Buscando archivos que contengan '{nombre_archivo_base}' en la carpeta ID: {carpeta_id}")
    
    results = service.files().list(
        q=query_sin_extension,
        pageSize=10,
        fields="files(id, name)"
    ).execute()
    
    items = results.get('files', [])

    if not items:
        print(f'No se encontró ningún archivo que contenga "{nombre_archivo_base}" en la carpeta especificada. Intentando con la extensión .xls y .xlsx.')
        
        # Si no se encontró ningún archivo, intentar con el nombre completo incluyendo las extensiones
        results_xls = service.files().list(
            q=query_xls,
            pageSize=10,
            fields="files(id, name)"
        ).execute()
        
        items.extend(results_xls.get('files', []))  # Agregar resultados de .xls

        results_xlsx = service.files().list(
            q=query_xlsx,
            pageSize=10,
            fields="files(id, name)"
        ).execute()
        
        items.extend(results_xlsx.get('files', []))  # Agregar resultados de .xlsx

    if not items:
        print(f'No se encontró ningún archivo que contenga "{nombre_archivo_base}" ni con ni sin extensión en la carpeta especificada.')
        
    else:
        # Mostrar coincidencias para ver posibles diferencias en el nombre del archivo
        print(f"Se encontraron {len(items)} archivos con nombres similares. Verificando coincidencias...")
        coincidencias = []
        for item in items:
            print(f"Nombre del archivo en Google Drive: {item['name']}")
            coincidencias.append(item)

        # Ordenar por la mayor similitud de nombre (puedes personalizar esta parte si deseas otro criterio)
        coincidencias_ordenadas = sorted(coincidencias, key=lambda x: x['name'])  # Ordenar alfabéticamente

        # Descargar el archivo más similar
        archivo_seleccionado = coincidencias_ordenadas[0]  # Tomar el primero como el más similar
        file_id = archivo_seleccionado['id']
        nombre_archivo_encontrado = archivo_seleccionado['name']
        
        request = service.files().get_media(fileId=file_id)
        destino_local = f"../TEMPORAL/{nombre_archivo_encontrado}"  # Actualizar la ruta según el nombre encontrado
        fh = io.FileIO(destino_local, 'wb')
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = downloader.next_chunk()
            print(f"Descargando {nombre_archivo_encontrado}: {int(status.progress() * 100)}%")

        print(f"Archivo {nombre_archivo_encontrado} descargado correctamente en {destino_local}.")

# Descargar archivos según las sucursales y fechas encontradas
for i in range(len(sucursal_columna)):
    sucursal = sucursal_columna[i]
    fecha = fecha_columna[i]

    # Verificar si la sucursal tiene una carpeta asignada
    if sucursal in carpetas_sucursales:
        carpeta_id = carpetas_sucursales[sucursal]
        nombre_archivo_base = f"{sucursal} {fecha}"  # Formato base para buscar, ignorando la extensión

        print(f"Descargando archivo que contenga '{nombre_archivo_base}' desde la carpeta {sucursal} (ID: {carpeta_id})")

        # Descargar archivo
        descargar_archivo(nombre_archivo_base, carpeta_id)
    else:
        print(f"Sucursal {sucursal} no tiene carpeta asignada.")
