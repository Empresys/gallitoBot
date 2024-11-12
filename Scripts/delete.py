import os
import shutil

# Ruta de la carpeta que deseas vaciar
carpeta = "../TEMPORAL"

# Eliminar todos los archivos en la carpeta
for archivo in os.listdir(carpeta):
    ruta_archivo = os.path.join(carpeta, archivo)
    try:
        if os.path.isfile(ruta_archivo) or os.path.islink(ruta_archivo):
            os.unlink(ruta_archivo)  # Eliminar archivos o enlaces simbólicos
        elif os.path.isdir(ruta_archivo):
            shutil.rmtree(ruta_archivo)  # Eliminar directorios y su contenido
    except Exception as e:
        print(f'No se pudo eliminar {ruta_archivo} debido a {e}')