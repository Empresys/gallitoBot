import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import difflib

# Definir los datos de la cuenta de correo para enviar notificaciones
EMAIL_SENDER = 'gallito2.empresys@gmail.com'
EMAIL_PASSWORD = 'ygvt fmgi clfj uccw'
EMAIL_RECEIVER = ''  # Se definirá más adelante

# Rutas de los archivos
ruta_archivo_excel = '../TEMPORAL/datos_factura.xlsx'
ruta_log = '../Log/Log.xlsx'  # Ruta del archivo de log

# Correo para copias
CC_EMAILS = ['gallito2.empresys@gmail.com', 'gallitomananerofacturas@gmail.com', 'soporte@empresys.com']

# Función para enviar un correo cuando un archivo no se encuentre o el monto no coincida
def enviar_correo_notificacion(asunto, mensaje):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SENDER
    msg['To'] = EMAIL_RECEIVER  # Se asignará más tarde
    msg['Cc'] = ', '.join(CC_EMAILS)  # Asignar los correos de copia
    msg['Subject'] = asunto

    # Agregar el cuerpo del mensaje
    msg.attach(MIMEText(mensaje, 'plain'))

    try:
        # Conectar al servidor SMTP de Gmail
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.sendmail(EMAIL_SENDER, [EMAIL_RECEIVER] + CC_EMAILS, msg.as_string())
        server.quit()
        print("Correo de notificación enviado correctamente.")
    except Exception as e:
        print(f"Error al enviar el correo de notificación: {e}")

# Leer el archivo de Excel original
try:
    df = pd.read_excel(ruta_archivo_excel)
    print("Archivo de Excel leído correctamente.")
except Exception as e:
    print(f"Error al leer el archivo de Excel: {e}")
    exit()

# Verificar que las columnas 'Sucursal', 'Fecha' y 'Consecutivo' existan en el archivo original
if 'Sucursal' not in df.columns or 'Fecha' not in df.columns or 'Folio' not in df.columns:
    print("Las columnas 'Sucursal', 'Fecha' o 'Folio' no se encuentran en el archivo de Excel.")
    exit()

# Obtener el destinatario desde la celda I2 (columna 8, fila 1)
EMAIL_RECEIVER = 'soporte@empresys.com'
print("Destinatario:", EMAIL_RECEIVER)

# Unir las columnas 'Fecha' y 'Sucursal' para formar el nombre del archivo
df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%d')  # Asegurarse de que la fecha esté en el formato correcto
df['Archivo'] = df['Sucursal'] + ' ' + df['Fecha']  # Crear el nombre del archivo basado en Sucursal y Fecha
df['Archivo'] = df['Archivo'].str.strip()

# Función para encontrar archivos en formato .xlsx o .xls
def obtener_archivo_formato(archivo_base):
    posibles_rutas = [
        f"../TEMPORAL/{archivo_base}.xlsx",
        f"../TEMPORAL/{archivo_base}.xls"
    ]
    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            return ruta
    return None

# Leer cada archivo con el nombre formado y realizar las validaciones
for index, row in df.iterrows():
    archivo_nombre = row['Archivo']
    folio = row['Folio']  # Obtener el folio
    monto_datos_factura = row['Monto']  # Supongamos que existe una columna "Monto" para la comparación

    # Buscar el archivo
    archivo_encontrado = obtener_archivo_formato(archivo_nombre)

    if archivo_encontrado:
        print(f"Archivo {archivo_encontrado} encontrado correctamente.")
        # Leer el archivo encontrado y buscar el numcheque que corresponda al Folio
        try:
            archivo_df = pd.read_excel(archivo_encontrado)
            if 'numcheque' in archivo_df.columns and 'total' in archivo_df.columns:
                fila_encontrada = archivo_df[archivo_df['numcheque'] == folio]
                if not fila_encontrada.empty:
                    monto_archivo = fila_encontrada['total'].values[0]
                    # Validar si los montos coinciden
                    if monto_datos_factura != monto_archivo:
                        print(f"El monto no coincide para el folio {folio}.")
                        enviar_correo_notificacion(
                            asunto=f"Discrepancia de Monto para Folio {folio}",
                            mensaje=(f"El monto del folio {folio}({monto_datos_factura}) "
                                     f"no coincide con el monto encontrado en el archivo verificado.")
                        )
                        # Actualizar log
                        log_df = pd.read_excel(ruta_log)
                        log_df.loc[log_df['Consecutivo'] == folio, 'Error'] = 'MONTO NO COINCIDE'
                        log_df.loc[log_df['Consecutivo'] == folio, 'Finalizado'] = 'ERROR'
                        log_df.to_excel(ruta_log, index=False)
                        # Eliminar archivo
                        os.remove(ruta_archivo_excel)
                        break
                else:
                    print(f"No se encontró el folio {folio} en el archivo verificado.")
                    enviar_correo_notificacion(
                        asunto=f"Folio {folio} No Encontrado",
                        mensaje=f"No se encontró el folio {folio} en el archivo '{archivo_nombre}'."
                    )
                    # Actualizar log
                    log_df = pd.read_excel(ruta_log)
                    log_df.loc[log_df['Consecutivo'] == folio, 'Error'] = 'FOLIO NO ENCONTRADO'
                    log_df.loc[log_df['Consecutivo'] == folio, 'Finalizado'] = 'ERROR'
                    log_df.to_excel(ruta_log, index=False)
                    # Eliminar archivo
                    os.remove(ruta_archivo_excel)
                    break
            else:
                print("El archivo verificado no tiene las columnas 'numcheque' o 'total'.")
        except Exception as e:
            print(f"Error al procesar el archivo {archivo_encontrado}: {e}")
    else:
        print(f"No se encontró el archivo: {archivo_nombre}. Enviando correo de notificación...")
        enviar_correo_notificacion(
            asunto=f"Archivo no encontrado para folio {folio}",
            mensaje=(f"No hemos encontrado el archivo correspondiente al folio {folio} "
                     f"({archivo_nombre}) en la carpeta especificada.")
        )
        # Actualizar log
        log_df = pd.read_excel(ruta_log)
        log_df.loc[log_df['Consecutivo'] == folio, 'Error'] = 'FOLIO NO ENCONTRADO'
        log_df.loc[log_df['Consecutivo'] == folio, 'Finalizado'] = 'ERROR'
        log_df.to_excel(ruta_log, index=False)
        # Eliminar archivo
        os.remove(ruta_archivo_excel)
        break
