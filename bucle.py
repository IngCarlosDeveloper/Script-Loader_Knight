import pyautogui
import pandas as pd
import subprocess
import time

ruta_excel = r"C:\\Users\\user\\Documents\\Trabajo\\Mercancia\\utilitarios\\LINK_KNIGHT_ULTRA.xlsm"

# Leer el archivo Excel y forzar que las columnas se lean como texto
datos = pd.read_excel(ruta_excel, sheet_name=0, dtype={'CODIGO': str, 'RECIBIDO': str})

# Reemplazar valores nulos con '0' en ambas columnas
datos['CODIGO'] = datos['CODIGO'].fillna('0')
datos['RECIBIDO'] = datos['RECIBIDO'].fillna('0')

time.sleep(5)

for index, fila in datos.iterrows():
    codigo = fila['CODIGO']
    cantidad = fila['RECIBIDO']

    if cantidad != "0":
        pyautogui.write(codigo)
        pyautogui.press('enter')

        pyautogui.write(cantidad)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')

        time.sleep(0.5)

stop = input("Presione ENTER para comenzar el proceso de limpieza de su sistema:")

subprocess.run(["python", r"C:\\Users\\user\\Desktop\\Loader_Knight\\limpieza.py"])
