import os
import subprocess
import time
import pyperclip
import win32com.client

subprocess.Popen(["C:\\a2Softway\\DataMigradorMR\\a2DataMigradorMR.exe"])

pyperclip.copy(r"C:\Users\user\Documents\Trabajo\Mercancia\utilitarios\zHoja3.txt")

time.sleep(1)

subprocess.run([r"C:\Users\user\Desktop\Loader_Knight\Tiny Task\exe\copy.exe"])

pyperclip.copy(r"C:\a2Softway\Empre001\Data")

subprocess.run([r"C:\Users\user\Desktop\Loader_Knight\Tiny Task\exe\datamigrador.exe"])

time.sleep(32)

os.system('taskkill /F /im a2DataMigradorMR.exe"')

subprocess.Popen([r"C:\\a2Softway\\dbsys\\dbsys.exe"])

time.sleep(2)

subprocess.run([r"C:\Users\user\Desktop\Loader_Knight\Tiny Task\exe\sql.exe"])

os.chdir(r"C:\a2Softway")

subprocess.Popen(["a2Admin.exe"])

time.sleep(1)

subprocess.run([r"C:\Users\user\Desktop\Loader_Knight\Tiny Task\exe\logeo.exe"])

time.sleep(2)

pyperclip.copy(r"C:\Users\user\Documents\Trabajo\Mercancia\utilitarios")

subprocess.run([r"C:\Users\user\Desktop\Loader_Knight\Tiny Task\exe\precio.exe"])

time.sleep(2)

# Ruta del archivo Excel
ruta_excel = r"C:\\Users\\user\\Documents\\Trabajo\\Mercancia\\utilitarios\\LINK_KNIGHT_ULTRA.xlsm"

# Crear una instancia de Excel
excel = win32com.client.Dispatch("Excel.Application")

# Opcional: hacer visible Excel (útil para desarrollo o depuración)
excel.Visible = True

# Abrir el archivo Excel
libro = excel.Workbooks.Open(ruta_excel)

try:
    # Seleccionar la Hoja 1
    hoja1 = libro.Sheets(1)  # Índice basado en 1 para Excel
    hoja1.Activate()

    print("Se abrió el archivo y se activó la Hoja 1.")

except Exception as e:
    print(f"Error al abrir el archivo o activar la Hoja 1: {e}")

subprocess.run([r"C:\\Users\\user\\Desktop\\Loader_Knight\\Tiny Task\\exe\\ordencompra.exe"])

time.sleep(1)

subprocess.run(["python", r"C:\\Users\\user\\Desktop\Loader_Knight\\bucle.py"])

