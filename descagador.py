import pyautogui
import pandas as pd
import subprocess
import time

#RECORDAR SIEMPRE CAMBIAR LA RUTA DEL EXCEL Y VERIFICAR EL NOMBRE DE LA COLUMNA DONDE VAYA LA CANTIDAD

time.sleep(5)

#CAMBIAR SIEMPRE LA RUTA DEL EXCEL ANTES DE EJECUTAR
#"C:\Users\user\Desktop\DESCARGO.xlsx"
ruta_excel = r"C:\Users\user\Desktop\DESCARGAR.xlsx"

# Leer el archivo Excel y forzar que las columnas se lean como texto
#VERIFICAR SIEMPRE EL NOMBRE DE LA COLUMNA DONDE ESTARAN LAS CANTIDADES
datos = pd.read_excel(ruta_excel, sheet_name=0, dtype={'CODIGO': str, 'ENVIADO': str})

# Reemplazar valores nulos con '0' en ambas columnas
datos['CODIGO'] = datos['CODIGO'].fillna('0')
datos['ENVIADO'] = datos['ENVIADO'].fillna('0')

for index, fila in datos.iterrows():
    codigo = fila['CODIGO']
    cantidad = fila['ENVIADO']

    if cantidad != "0":
        pyautogui.write(codigo)
        pyautogui.press('enter')

        pyautogui.write(cantidad)
        pyautogui.press('enter')

        time.sleep(0.5)

stop = input("Presione ENTER para comenzar el proceso de limpieza de su sistema:")

subprocess.run(["python", r"C:\\Users\\user\\Desktop\\Loader_Knight\\limpieza.py"])
