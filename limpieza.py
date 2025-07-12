import os
import subprocess
import time
import win32com.client

ruta_excel = r"C:\\Users\\user\\Documents\\Trabajo\\Mercancia\\utilitarios\\LINK_KNIGHT_ULTRA.xlsm"

# Crear una instancia de Excel
excel = win32com.client.Dispatch("Excel.Application")

# Abrir el archivo Excel
libro = excel.Workbooks.Open(ruta_excel)

try:
    # Ejecutar las macros en el orden deseado
    excel.Application.Run("LINK_KNIGHT_ULTRA.xlsm!DELETE")

    print("Todas las macros se ejecutaron con éxito.")

except Exception as e:
    print(f"Error al ejecutar las macros: {e}")

finally:
    # Guardar y cerrar el archivo
    libro.Close(SaveChanges=True)
    excel.Application.Quit()

subprocess.run([r"C:\\limpieza.bat"])

input("Presione ENTER cuando haya devuelto la data para seguir con los procesos: ")
