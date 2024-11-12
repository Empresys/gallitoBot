import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# Leer los archivos Excel y de texto
errores_df = pd.read_excel('../Folder/GallitoErrores.xlsx')  # Ruta del archivo de errores
factura_df = pd.read_excel('../TEMPORAL/datos_factura.xlsx')  # Ruta del archivo de facturas

# Verificar si el DataFrame de facturas tiene datos
if factura_df.empty:
    print("El DataFrame no tiene datos.")
    exit()

# Obtener el folio de la primera fila
folio = factura_df['Folio'].iloc[0]
print("Folio:", folio)

# Obtener el destinatario desde la celda I2 (columna 8, fila 1)
destinatario = factura_df.iloc[0, 8]  # Columna I (8) y fila 1 (0 en índice base 0)
print("Destinatario:", destinatario)

# Leer el mensaje de error desde el archivo error.txt con la codificación correcta
with open('../TEMPORAL/error.txt', 'r', encoding='utf-16') as file:
    mensaje_error = file.read()

# Imprimir el mensaje de error completo para depuración
print(f"Mensaje de error completo: {mensaje_error}")

# Inicializar variable para la descripción del error
descripcion_error = ""

# Extraer la descripción del error desde el mensaje de error después de la palabra "Error:"
if "Error:" in mensaje_error:
    # Dividir el mensaje por líneas y buscar la línea que contiene "Error:"
    lineas = mensaje_error.splitlines()
    for linea in lineas:
        if "Error:" in linea:
            error_msg = linea.split("Error:")[1].strip()  # Obtiene el texto después de "Error:"
            if "RegimenFiscalR" in error_msg:
                regimen_fiscal = factura_df.iloc[0, 4]  # Columna E (4) y fila 1 (0 en índice base 0)
                descripcion_error = f'El régimen fiscal {regimen_fiscal} es incorrecto.'
            elif "Nombre del receptor" in error_msg:
                nombre_receptor = factura_df.iloc[0, 1]  # Asumido: columna B (1) y fila 1 (0 en índice base 0)
                descripcion_error = f'La razón social {nombre_receptor} no coincide con los registros del SAT.'
            elif "RFC" in error_msg:
                rfc = factura_df.iloc[0, 2]  # Asumido: columna C (2) y fila 1 (0 en índice base 0)
                descripcion_error = f'El RFC {rfc} no existe en la lista de RFC inscritos no cancelados del SAT.'
            else:
                # Para cualquier otro error, simplemente capturamos el mensaje
                descripcion_error = error_msg
            break  # Salir del bucle una vez encontrado

# Imprimir la descripción del error para depuración
print(f"Descripción del error encontrada: {descripcion_error}")

# Crear cuerpo del correo con el folio y el detalle del error
def crear_cuerpo_error(folio, detalle):
    return f"""
Estimado cliente,

Te informamos que tras verificar la información proporcionada para la generación de tu factura con folio {folio}, encontramos los siguientes errores:

{detalle}

Favor de verificar la información e intentar nuevamente. https://factgallito.empresys.com
Saludos cordiales.
"""

# Crear el cuerpo del mensaje basado en el error detectado
cuerpo = crear_cuerpo_error(folio, f'Descripción del error: {descripcion_error}' if descripcion_error else "No se encontró una descripción del error específica.")

# Configurar el mensaje de correo
msg = MIMEMultipart()
msg['From'] = 'gallito2.empresys@gmail.com'  # Usar credenciales de gallito2
msg['To'] = destinatario  # Enviar a la dirección obtenida de la celda I2
msg['Cc'] = 'gallito2.empresys@gmail.com,gallitomananerofacturas@gmail.com,soporte@empresys.com'  # Copias a los otros destinatarios
msg['Subject'] = f'Error en la factura con folio {folio}'
msg.attach(MIMEText(cuerpo, 'plain'))

# Enviar el correo
try:
    with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
        servidor.starttls()
        servidor.login('gallito2.empresys@gmail.com', 'ygvt fmgi clfj uccw')  # Usa las credenciales de gallito2
        servidor.sendmail(msg['From'], [msg['To']] + msg['Cc'].split(','), msg.as_string())
    print(f'Correo enviado exitosamente a {destinatario}.')
except smtplib.SMTPAuthenticationError:
    print('Error de autenticación. Verifica tus credenciales de Gmail.')
except Exception as e:
    print(f'Error al enviar el correo: {e}')

# Actualizar el archivo de log
ruta_log = '../Log/Log.xlsx'  # Ruta del archivo de log

try:
    log_df = pd.read_excel(ruta_log)

    # Verificar si la columna 'Consecutivo' existe
    if 'Consecutivo' in log_df.columns:
        # Marcar el error en el archivo Log.xlsx
        log_df.loc[log_df['Consecutivo'] == folio, 'Error'] = 'ERRO EN LA FACTURA'
        log_df.loc[log_df['Consecutivo'] == folio, 'Finalizado'] = 'ERROR'

        # Guardar los cambios en el archivo de log
        log_df.to_excel(ruta_log, index=False)
        print(f"Se actualizó el log para el folio {folio}.")
    else:
        print("La columna 'Consecutivo' no se encuentra en el archivo de log.")

except Exception as e:
    print(f"Error al intentar actualizar el archivo de log: {e}")
