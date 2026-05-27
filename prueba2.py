import subprocess
import time
import pyautogui
import os
import pyperclip
import win32clipboard
import shutil
import pdfplumber
import re
import pandas as pd
import win32com.client
from openpyxl import load_workbook
import fnmatch
import requests
from datetime import date
import keyboard
import pydirectinput

# base_dir = os.path.dirname(__file__) #Esto extrae la ruta actual

# def click_en(x,y):
#     pyautogui.moveTo(x, y, duration=1)
#     pyautogui.click()

# #--------------------------------------PRESIONAR TECLAS ESPECIALES-------------------------------------------

# def presionar_teclas(*teclas):
#     pyautogui.hotkey(*teclas)


# def pegar_archivos_portapapeles(destino):

#     win32clipboard.OpenClipboard()

#     try:
#         archivos = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
#     except TypeError:
#         print("No hay archivos en el portapapeles")
#         win32clipboard.CloseClipboard()
#         return

#     win32clipboard.CloseClipboard()

#     for archivo in archivos:

#         nombre = os.path.basename(archivo)
#         destino_final = os.path.join(destino, nombre)

#         shutil.copy2(archivo, destino_final)

#         print(f"Copiado: {nombre}")

# def extraer_data():

#     exe = r"tiny_task\\exe"
#     ruta_tiny_task = os.path.join(base_dir, exe)

#     click_en(781, 23)
#     click_en(842, 706) #Clicks para despertar la ventana de anydesk

#     subprocess.run([f"{ruta_tiny_task}\\winR.exe"])

#     #pyperclip.copy(r"C:\procesador_winrar.bat")

#     pyautogui.write(r"C:\procesador.bat", interval=0.1) #Escribimos la ruta del .bat que esta en la otra pc
#     presionar_teclas("enter")
    
#     time.sleep(20)

#     subprocess.run([r"C:\open.bat"])

#     pegar_archivos_portapapeles(r"C:\a2Softway\Empre001")

#     os.makedirs(r"C:\a2Softway\Empre001\Data", exist_ok=True)

#     subprocess.run([
#         r"C:\Program Files\WinRAR\WinRAR.exe",
#         "x",
#         "-y",
#         r"C:\a2Softway\Empre001\zData.rar",
#         r"C:\a2Softway\Empre001\Data"
#     ])

#     print("Extracción completada")

#     shutil.copytree(
#         r"C:\a2Softway\Empre001\Data",
#         r"C:\a2DEMO\Empre001\Data",
#         dirs_exist_ok=True
#     )


# #extraer_data()

# time.sleep(3)

# pyautogui.write(r"C:\procesador.bat", interval=0.1)


time.sleep(5)

ruta = r"c:/comprimidor.bat"
pyperclip.copy(ruta)

pydirectinput.press('backspace')
# pydirectinput.keyDown('ctrl')
# pydirectinput.press('v')
# pydirectinput.keyUp('ctrl')
# pydirectinput.press("enter")