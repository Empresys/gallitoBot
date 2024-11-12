import pandas as pd

# Leer los archivos Excel
factura_df = pd.read_excel('../TEMPORAL/datos_factura.xlsx')  # Ruta del archivo de facturas

# Verificar si el DataFrame de facturas tiene datos
if factura_df.empty:
    print("El DataFrame no tiene datos.")
    exit()

# Obtener el folio de la primera fila
folio = factura_df['Folio'].iloc[0]
print("Folio:", folio)

# Ruta del archivo de log
ruta_log = '../Log/Log.xlsx'

try:
    # Leer el archivo de log
    log_df = pd.read_excel(ruta_log)

    # Verificar si la columna 'Consecutivo' existe
    if 'Consecutivo' in log_df.columns:
        # Marcar el folio como exitoso en el archivo Log.xlsx
        log_df.loc[log_df['Consecutivo'] == folio, 'Error'] = ''  # Dejar vacío el campo de error
        log_df.loc[log_df['Consecutivo'] == folio, 'Finalizado'] = 'Exitoso'  # Marcar como exitoso

        # Guardar los cambios en el archivo de log
        log_df.to_excel(ruta_log, index=False)
        print(f"Se actualizó el log para el folio {folio} como 'Exitoso'.")
    else:
        print("La columna 'Consecutivo' no se encuentra en el archivo de log.")

except Exception as e:
    print(f"Error al intentar actualizar el archivo de log: {e}")
