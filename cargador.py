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
import win32clipboard
import win32con
from openpyxl import load_workbook
import fnmatch
import requests
from datetime import date
import pydirectinput
import traceback
import threading
import keyboard

anydesk = 0
password = 0

#ruta actual
base_dir = os.path.dirname(__file__) #Esto extrae la ruta actual

correo_tiendas = {
    "nombre_tienda" : "correodeprueba@gmail.com",
}

libro_notas_entrega = {
    
}
#"00000001" : "DE CONTADO",
#"00000002" : "DE CONTADO"

direcciones_anydesk = {
    "Nombre_tienda": "direccion_anydesk",
}

lista_marcas = [
    "PENGUIN",
    "DKNY",
    "PURO EGO",
    "COMFY"
]

anydesk_password = "password del anydesk"

#-------------------------------------------------FUNCION PARA PAUSAR EL CODIGO-------------------------

pause_event = threading.Event()
pause_event.set()  # inicia sin pausa

def escuchar_pausa():
    while True:
        keyboard.wait('|')  # tecla Pause/Break del teclado
        if pause_event.is_set():
            print("⏸ PAUSADO")
            pause_event.clear()
        else:
            print("▶ REANUDADO")
            pause_event.set()

def check_pause():
    pause_event.wait()


#--------------------------------------FUNCION PARA TRAER LA RUTA DE LA IMAGEN----------------------------

def img(nombre):
    return str(os.path.join(base_dir, "imagenes_pyautogui", nombre))

#--------------------------------------HACER CLICKS EN COORDENADAS ESPECIFICAS----------------------------

def click_en(x,y):
    pyautogui.moveTo(x, y, duration=0.2)
    pyautogui.click()

#--------------------------------------PRESIONAR TECLAS ESPECIALES-------------------------------------------

def presionar_teclas(*teclas):
    pyautogui.hotkey(*teclas)

#--------------------------------------EXTRAER LA TASA DEL BCV-----------------------------------------------

def tasa_bcv():

    url = "https://dolarflow.com/api/oficial"

    r = requests.get(url)
    data = r.json()

    tasa = data["precio"]

    return tasa

# --------------------------------------VERIFICAR SI ANYDESK PIDE CONTRASEÑA----------------------------

def verificar_password(contra):

    #Recuerda verificar que hace si no hay contrasena
    inicio = time.time()

    while time.time() - inicio < 10:

        try:
            pos = pyautogui.locateOnScreen(img("contrasena_anydesk.png"), confidence=0.8)

            if pos:
                print("Campo de contrasena detectado")

                time.sleep(2)
                pyautogui.click(pos)
                pyautogui.write(contra, interval=0.1)
                pyautogui.press("enter")

                return True
        except:
            pass
        time.sleep(0.5)
        check_pause()

    print("No se solicito password")
    return False

#--------------------------------------------------LEVANTAR ANYDESK---------------------------------------------

def conectar_anydesk(id_anydesk , contra):

    #Abrir anydesk y conectarse directamente a una PC
    ruta = r"C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe " + str(id_anydesk)

    #print(ruta)
    subprocess.Popen(ruta)

    time.sleep(3)

    verificar_password(contra)

    for g in range(10):
        check_pause()
        time.sleep(1)

#--------------------------------------------------ACCEDER AL PORTAPAPELES---------------------------------------------

def pegar_archivos_portapapeles(destino):

    # win32clipboard.OpenClipboard()

    # try:
    #     archivos = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
    # except TypeError:
    #     print("No hay archivos en el portapapeles")
    #     win32clipboard.CloseClipboard()
    #     return

    # win32clipboard.CloseClipboard()

    # for archivo in archivos:

    #     nombre = os.path.basename(archivo)
    #     destino_final = os.path.join(destino, nombre)

    #     shutil.copy2(archivo, destino_final)

    #     print(f"Copiado: {nombre}")

    os.startfile(destino)

    time.sleep(2)

    pydirectinput.keyDown('ctrl')
    pydirectinput.press('v')
    pydirectinput.keyUp('ctrl')

    for _ in range(60):
        check_pause()
        time.sleep(1)


#--------------------------------------COPIAR ARCHIVO EN EL PORTAPAPELES------------------------------
def copiar_archivo_portapapeles(ruta_archivo):

    ruta_archivo = os.path.abspath(ruta_archivo)

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()

    data = ('\0'.join([ruta_archivo]) + '\0').encode('utf-16le')

    win32clipboard.SetClipboardData(win32con.CF_HDROP, data)

    win32clipboard.CloseClipboard()

    print("Archivo copiado al portapapeles")
    
#--------------------------------------------------INVOCAR EL .BAT PARA COMPRIMIR Y EXTRAER LA DATA---------------------------------------------
#Ya con el anydesk abierto se trae la data del a2 remoto y lo instala listo en el local
def extraer_data():

    check_pause()
    exe = r"tiny_task\\exe"
    ruta_tiny_task = os.path.join(base_dir, exe)

    click_en(781, 23)
    click_en(842, 706) #Clicks para despertar la ventana de anydesk

    subprocess.run([f"{ruta_tiny_task}\\winR.exe"])

    #pyperclip.copy(r"C:\procesador_winrar.bat")

    #pyautogui.write(r"C:\procesador.bat", interval=0.1) #Escribimos la ruta del .bat que esta en la otra pc
    #presionar_teclas("enter")

    # pydirectinput.write(r"c:\procesador.bat", interval=0.05)
    # pydirectinput.press("enter")

    ruta = r"c:/comprimidor.bat"
    pyperclip.copy(ruta)

    pydirectinput.press('backspace')
    pydirectinput.keyDown('ctrl')
    pydirectinput.press('v')
    pydirectinput.keyUp('ctrl')
    pydirectinput.press("enter")
    #Ya aqui ejecutamos el .bat para comprimir la data
    for _ in range(40):
        check_pause()
        time.sleep(1)

    subprocess.run([r"C:\open.bat"])

    pegar_archivos_portapapeles(r"C:\a2Softway\Empre001")

    os.makedirs(r"C:\a2Softway\Empre001\Data", exist_ok=True)

    subprocess.run([
        r"C:\Program Files\WinRAR\WinRAR.exe",
        "x",
        "-y",
        r"C:\a2Softway\Empre001\zData.rar",
        r"C:\a2Softway\Empre001\Data"
    ])

    print("Extracción completada")

    shutil.copytree(
        r"C:\a2Softway\Empre001\Data",
        r"C:\a2DEMO\Empre001\Data",
        dirs_exist_ok=True
    )

#------------------------------------------LEER EL PDF Y EXTRAER DATOS------------------------------------------

def pdf_reader(nombre_pdf):

    with pdfplumber.open(nombre_pdf) as pdf:
        pagina = pdf.pages[0] # 0 = 1era Hoja
        texto = pagina.extract_text()

    match = re.search(r"Cliente:\s*(.*)", texto)
    match_id = re.search(r"CENTRO EMPRESARIAL STURGIS\S*(.*)", texto)

    #Bloque para extraer el nombre del cliente
    if match:
        cliente = match.group(1).strip()
        print(cliente)

    if match_id:
        transaccion = match_id.group(1).strip()
        print(transaccion)


    # lineas = texto.split("\n")

    # numero_transaccion = None

    # for i, linea in enumerate(lineas):
    #     if "NOTA DE ENTREGA" in linea.upper():
    #         if i + 1 < len(lineas):
    #             posible = lineas[i+1].strip()

    #             if re.fullmatch(r"\d{8}", posible):
    #                 numero_transaccion = posible
    #                 break

    # print(numero_transaccion)

    return cliente, transaccion

    #aqui llamamos a la funcion para procesar las notas de entrega

#------------------------------------------COMBINAR EXCELS------------------------------------------------

#esta funcion combina el excel que suelta el smarttools y lo pasa a un formato adecuado para que el personal pueda revisar
def combinar_excels(excel_origen):
    
    tmp = str(excel_origen) + ".xls" #esta variable es el id de la nota mas el .xlsx
    archivo_A = os.path.join(base_dir, "Nota_Entrega", tmp)
    plantilla = os.path.join(base_dir, "plantilla", "formato-mercancia.xlsx")

    df_A = pd.read_excel(archivo_A, header=None, dtype=str, engine="xlrd")

    df_A.columns = [
        "CODIGO","REFERENCIA","DESCRIPCION","TALLA/MODELO",
        "COLOR/CATEGORIA","CANTIDAD","MONEDA","PRECIO_BS",
        "DEPARTAMENTO","COSTO_ACTUAL","PRECIO_DIV"
    ]

    df_A["PRECIO_DIV"] = (
        df_A["PRECIO_DIV"]
        .astype(str)
        .str.replace(",",".", regex=False)
        .astype(float) * 1.16
    )

    mapa_columnas = {
        "CODIGO": "CODIGO",
        "REFERENCIA": "REFERENCIA",
        "DESCRIPCION": "DESCRIPCION",
        "TALLA/MODELO": "TALLA/MODELO",
        "COLOR/CATEGORIA": "COLOR/CATEGORIA",
        "ENVIADO": "CANTIDAD",
        "PRECIO": "PRECIO_DIV",
        "DEPARTAMENTO": "DEPARTAMENTO"
    }

    # abrir plantilla
    wb = load_workbook(plantilla)
    ws = wb.active

    # leer encabezados de plantilla
    encabezados = {}
    for col in range(1, ws.max_column + 1):
        encabezados[ws.cell(row=1, column=col).value] = col

    fila_inicio = 2

    for i, fila in df_A.iterrows():

        for col_destino, col_origen in mapa_columnas.items():

            col_excel = encabezados[col_destino]

            celda = ws.cell(
                row=fila_inicio + i,
                column=col_excel,
                value=fila[col_origen]
            )

            if col_destino == "CODIGO":
                celda.number_format = "@"

    tienda_actual = libro_notas_entrega[excel_origen]

    hoy = date.today().strftime(f"%d-%m-%Y")

    nombre_excel = f"CARGAR MERCANCIA {tienda_actual} {hoy} {excel_origen}.xlsx" 
    ruta_guardar = os.path.join(base_dir, "Data")

    wb.save(f"{ruta_guardar}\\{nombre_excel}")

    print("Excel generado conservando formato y formulas")

    destino = os.path.join(base_dir, "Nota_Entrega", "completados")

    shutil.move(archivo_A, destino)

#------------------------------------------DESCARGA LOS PDFS DEL CORREO QUE VIENEN DEL DEPOSITO------------------------------------------------

def leer_pdf_correo():
    #En la carpeta Data es en donde guardaremos los pdfs
    output_folder = os.path.join(base_dir, "pdf")
    os.makedirs(output_folder, exist_ok=True)

    #conectar con outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    #Seleccionar la bandeja de entrada
    inbox = outlook.GetDefaultFolder(6) # 6 representa la bandeja de entrada
    processed_folder_name = "Nota_Entrega"
    processed_folder = inbox.Folders(processed_folder_name)

    #Iterar a traves de los correos electronicos
    messages = inbox.Items
    for message in messages:
        try:

            subject = message.Subject #Titulo del correo
            sender = message.SenderName #Remitente

            print(f"Título: {subject} | Remitente: {sender}")

            if fnmatch.fnmatch(subject, "*ENTREGA*"):

                #Descargar y renombrar archivos adjuntos:
                if message.Attachments.Count > 0:
                    for attachment in message.Attachments:

                        #Crear un nombre de archivo basado en el asunto del correo
                        filename = f"{subject} - {attachment.FileName}".replace(":", "-").replace("/", "-")
                        #filename = attachment.FileName.replace(":", "-").replace("/", "-")
                        #filename = f"{attachment.FileName} - {subject}".replace(":", "-").replace("/", "-")
                        file_path = os.path.join(output_folder, filename)

                        # Evitar sobrescritura
                        counter = 1
                        original_path = file_path
                        while os.path.exists(file_path):
                            base, ext = os.path.splitext(original_path)
                            file_path = f"{base} ({counter}){ext}"
                            counter += 1

                        # Guardar el archivo adjunto
                        attachment.SaveAsFile(file_path)
                        print(f"Archivo descargado: {file_path}")

                        nombre_tienda, id_transaccion = pdf_reader(file_path)

                        libro_notas_entrega[id_transaccion] = nombre_tienda

                # Mover el correo procesado a la carpeta "Completados"
                message.Move(processed_folder)
                print(f"Correo movido a la carpeta '{processed_folder_name}'")


            else:
                print("No se encontraron NOTAS DE ENTREGA.")

        except Exception as e:
            print(f"Error procesando un correo: {e}")

    #aqui llamas a las funciones para extraer las notas de entrega

#------------------------------ENVIA LOS CORREOS CON EL EXCEL PARA QUE LAS TIENDAS REVISEN---------------------

def envia_correo():

    carpeta_excel = os.path.join(base_dir, "Data")

    outlook = win32com.client.Dispatch("Outlook.Application")

    for archivo in os.listdir(carpeta_excel): #Enlista todos los excels dentro de la carpeta Data

        if archivo.lower().endswith(".xlsx"):

            ruta_archivo = os.path.join(carpeta_excel, archivo)

            nombre_archivo = archivo.upper()

            for tienda, correo in correo_tiendas.items():

                if tienda in nombre_archivo:

                    mail = outlook.CreateItem(0)

                    mail.To = correo
                    mail.Subject = f"Nota de Entrega - {tienda}"
                    mail.Body = f"""
 Buen día,
 Adjunto encontrará el archivo correspondiente al envio de mercancia a la tienda {tienda}.

 Saludos.

 Nota: No responder a este correo.
 """
#                     mail.Body = f"""
# Buen dia, correo de prueba, si lees esto, por favor ignoralo, gracias :D.
# """
#                     mail.Attachments.Add(ruta_archivo)

#                     mail.Send()

#                     print(f"Correo enviado a {tienda} ({correo})")

#                     destino = os.path.join(carpeta_excel, "completados")
#                     shutil.move(ruta_archivo, destino)

                    break

#-------------------------------------------CONVERTIR XLSX EN XLS-----------------------------------------

def convertir_xls(ruta_xlsx):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(ruta_xlsx)

    ws = wb.ActiveSheet

    for col in range(1, ws.UsedRange.Columns.Count + 1):
        if ws.Cells(1, col).Value == "CODIGO":
            
            # Aplicar formato TEXTO a toda la columna
            ws.Columns(col).NumberFormat = "@"
            break

    ruta_xls = ruta_xlsx.replace(".xlsx", ".xls")

    wb.SaveAs(ruta_xls, FileFormat=56)
    wb.Close()

    excel.Quit()

    return ruta_xls

#---------------------------------------PREPARAR EXCEL PARA INSERTAR EN SMARTOOLS------------------------

def excel_smarttools(ruta_excel, nombre_excel):

    columnas_cero = [
        "REFERENCIA",
        "DESCRIPCION",
        "TALLA/MODELO",
        "COLOR/CATEGORIA",
        "ENVIADO",
        "RECIBIDO",
    ]

    nombre_excel = nombre_excel.lower()

    excel = pd.read_excel(ruta_excel, dtype={"CODIGO": str})

    excel[columnas_cero] = excel[columnas_cero].fillna(0)

    excel["DESCRIPCION"] = (
    excel["DESCRIPCION"].fillna("").astype(str) +
    " T/" +
    excel["TALLA/MODELO"].fillna("").astype(str) +
    " " +
    excel["COLOR/CATEGORIA"].fillna("").astype(str)
    )

    excel["PRECIO"] = (excel["PRECIO"]
                       .astype(str)
                       .str.replace(",",".", regex=False)
                       .astype(float)
                       )
    
    excel["PVP_SIN_IVA"] = excel["PRECIO"] / 1.16

    excel["MONEDA"] = 2

    excel["IVA"] = 16

    excel["COSTO"] = excel["PVP_SIN_IVA"] / 2

    tasa_dolar = tasa_bcv()

    excel["COSTO_BS"] = excel["COSTO"] * tasa_dolar

    excel.loc[0,"DEPOSITO"] = 1

    excel.loc[0,"PROVEEDOR"] = 1

    excel.loc[0,"FECHA"] = date.today().strftime(f"%d/%m/%Y")

    orden_columnas = [
        "CODIGO",
        "REFERENCIA",
        "DESCRIPCION",
        "RECIBIDO",
        "PRECIO",
        "PVP_SIN_IVA",
        "DEPARTAMENTO",
        "MONEDA",
        "IVA",
        "COSTO",
        "COSTO_BS",
        "TALLA/MODELO",
        "COLOR/CATEGORIA",
        "DEPOSITO",
        "PROVEEDOR",
        "FECHA"
    ]

    excel = excel[orden_columnas]

    #Guardar el dataframe completo

    ruta_nuevo_excel = os.path.join(base_dir, "excel_cargar", "tmp", nombre_excel)

    #ruta_nuevo_excel = ruta_nuevo_excel + nombre_excel

    nombre_xls = nombre_excel.replace(".xlsx", ".xls")

    excel.to_excel(ruta_nuevo_excel, index=False)
    print(f"Archivo {nombre_excel}.xlsx generado exitosamente")

    xls_tmp = convertir_xls(ruta_nuevo_excel) #aqui estaria la ruta del .xls temporal

    destino1 = os.path.join(base_dir, "excel_cargar", "terminados")
    shutil.move(ruta_excel, destino1)



    cargar_smarttools_sin_costo(nombre_xls)

#------------------------------------INGRESAR EL EXCEL EN EL SMARTTOOLS SIN COSTO-----------------------------

def cargar_smarttools_costo(nombre_excel):

    check_pause()
    exe = r"tiny_task\\exe"
    ruta_tiny_task = os.path.join(base_dir, exe)

    #iniciamos smartools

    os.chdir(r"C:\a2DEMO")

    subprocess.run([r"SmartTools.exe"])

    time.sleep(3) #esperamos unos minutos a que abra el smarttools

    pyautogui.write("MASTER", interval=0.1)
    presionar_teclas("enter")
    pyautogui.write("password", interval=0.1)
    presionar_teclas("enter")
    presionar_teclas("enter")

    check_pause()
    subprocess.run([f"{ruta_tiny_task}\\1-acceder-cargar.exe"])

    time.sleep(2)

    #aqui tenemos que poner la ruta del archivo sin el nombre del mismo
    pyautogui.write(r"C:\Users\user\Desktop\BOTS\Cargador_Smartools\masivo")
    presionar_teclas("enter")
    
    check_pause()
    subprocess.run([f"{ruta_tiny_task}\\2-nombre-archivo.exe"])

    time.sleep(2)

    pyautogui.write(nombre_excel)

    presionar_teclas("enter")
    presionar_teclas("enter")

    subprocess.run([f"{ruta_tiny_task}\\3-cargar-perfil-costo.exe"])
    print("esperando mientras se ingresa la mercancia")
    for _ in range(60):
        check_pause()
        time.sleep(1)

    check_pause()
    subprocess.run([f"{ruta_tiny_task}\\6-salir.exe"])

    time.sleep(2)
    check_pause()
    subprocess.run([r"C:\comprimidor.bat"])
    check_pause()

    #copiar_archivo_portapapeles(r"C:\a2DEMO\Empre001\xDatafinal.rar")
    subprocess.run([r"C:\copiar.bat"])
    check_pause()
    #click_en(337, 736) #Anydesk en la barra de tarea
    click_en(1016, 369) #carpeta para pegar

    pydirectinput.keyDown('ctrl')
    pydirectinput.press('v')
    pydirectinput.keyUp('ctrl')

    #presionar_teclas("ctrl", "v")

    for _ in range(60):
        check_pause()
        time.sleep(1)

    subprocess.run([f"{ruta_tiny_task}\\winR.exe"])

    # pyautogui.write(r"C:\descomprimidor.bat", interval=0.1) #Escribimos la ruta del .bat que esta en la otra pc
    # presionar_teclas("enter")

    ruta = r"c:\descomprimidor.bat"
    pyperclip.copy(ruta)


    pydirectinput.keyDown('ctrl')
    pydirectinput.press('v')
    pydirectinput.keyUp('ctrl')
    pydirectinput.press("enter")

    time.sleep(10)

    check_pause()
    os.system('taskkill /F /im AnyDesk.exe"')

    os.system('taskkill /F /im explorer.exe"')
    time.sleep(0.5)
    os.system("start explorer.exe")

    xd = input("Presione ENTER para continuar con el proceso")

#------------------------------------INGRESAR EL EXCEL EN EL SMARTTOOLS SIN COSTO-----------------------------

def cargar_smarttools_sin_costo(nombre_excel):

    exe = r"tiny_task\\exe"
    ruta_tiny_task = os.path.join(base_dir, exe)

    #iniciamos smartools

    os.chdir(r"C:\a2DEMO")

    subprocess.run([r"SmartTools.exe"])
    check_pause()

    time.sleep(3) #esperamos unos segundos a que abra el smarttools

    pyautogui.write("MASTER", interval=0.1)
    presionar_teclas("enter")
    pyautogui.write("password", interval=0.1)
    presionar_teclas("enter")
    presionar_teclas("enter")

    subprocess.run([f"{ruta_tiny_task}\\1-acceder-cargar.exe"])

    time.sleep(2)

    #aqui tenemos que poner la ruta del archivo sin el nombre del mismo
    #pyautogui.write(r"C:\Users\user\Desktop\BOTS\Cargador_Smartools\Data")
    pyautogui.write(r"C:\Users\user\Desktop\BOTS\Cargador_Smartools\excel_cargar\tmp")
    presionar_teclas("enter")
    
    subprocess.run([f"{ruta_tiny_task}\\2-nombre-archivo.exe"])

    time.sleep(2)

    pyautogui.write(nombre_excel)

    presionar_teclas("enter")
    presionar_teclas("enter")

    subprocess.run([f"{ruta_tiny_task}\\3-2-cargar-perfil-sin-costo.exe"])
    print("esperando mientras se ingresa la mercancia")
    for _ in range(60):
        check_pause()
        time.sleep(1)




    check_pause()
    subprocess.run([f"{ruta_tiny_task}\\3-3-Ir-importar-transaccion.exe"])

    time.sleep(2)

    #aqui tenemos que poner la ruta del archivo sin el nombre del mismo
    #pyautogui.write(r"C:\Users\user\Desktop\BOTS\Cargador_Smartools\Data")
    pyautogui.write(r"C:\Users\user\Desktop\BOTS\Cargador_Smartools\excel_cargar\tmp")
    presionar_teclas("enter")
    check_pause()
    subprocess.run([f"{ruta_tiny_task}\\2-nombre-archivo.exe"])

    time.sleep(2)

    pyautogui.write(nombre_excel)

    presionar_teclas("enter")
    presionar_teclas("enter")
    check_pause()
    subprocess.run([f"{ruta_tiny_task}\\4-cargar-compra-mercancia"])

    for _ in range(20):
        check_pause()
        time.sleep(1)

    subprocess.run([f"{ruta_tiny_task}\\4-salir-actualizar-datos.exe"])

    time.sleep(5)
    check_pause()
    subprocess.run([r"C:\comprimidor.bat"])

    time.sleep(20)

    #copiar_archivo_portapapeles(r"C:\a2DEMO\Empre001\xDatafinal.rar")
    subprocess.run([r"C:\copiar.bat"])
    check_pause()
    #click_en(337, 736) #Anydesk en la barra de tarea
    click_en(1016, 369) #carpeta para pegar

    #Aqui pegamos el XdataFinal en la otra maquina
    pydirectinput.keyDown('ctrl')
    pydirectinput.press('v')
    pydirectinput.keyUp('ctrl')

    for _ in range(60):
        check_pause()
        time.sleep(1)

    subprocess.run([f"{ruta_tiny_task}\\winR.exe"])

    # pyautogui.write(r"C:\descomprimidor.bat", interval=0.1) #Escribimos la ruta del .bat que esta en la otra pc
    # presionar_teclas("enter")

    ruta = r"c:\descomprimidor.bat"
    pyperclip.copy(ruta)
    check_pause()

    pydirectinput.keyDown('ctrl')
    pydirectinput.press('v')
    pydirectinput.keyUp('ctrl')
    pydirectinput.press("enter")    

    time.sleep(10)

    #os.system('taskkill /F /im AnyDesk.exe"')

    #time.sleep(10)

    #os.system('taskkill /F /im explorer.exe"')
    #time.sleep(0.5)
    #os.system("start explorer.exe")

    xd = input("Presione ENTER para continuar con el proceso")

    origen = os.path.join(base_dir, "excel_cargar", "tmp", nombre_excel)
    destino = os.path.join(base_dir, "excel_cargar", "tmp", "completados")
    shutil.move(origen, destino)

#------------------------------------LEVANTAR SMARTTOOLS Y ACCEDER A EXPORTAR-------------------------------
#Esta funcion solo accede, no exporta, eso lo hace la funcion exportar_notas_entrega, ya que ese proceso puede ser en bucle
def acceder_extraer_nota_entrega():
    
    exe = r"tiny_task\\exe"
    ruta_tiny_task = os.path.join(base_dir, exe)

    #iniciamos smartools

    os.chdir(r"C:\a2DEMO")

    subprocess.run([r"SmartTools.exe"])

    time.sleep(15) #esperamos unos segundos a que abra el smarttools

    pyautogui.write("MASTER", interval=0.1)
    presionar_teclas("enter")
    pyautogui.write("password", interval=0.1)
    presionar_teclas("enter")
    presionar_teclas("enter")

    subprocess.run([f"{ruta_tiny_task}\\1-NE-ACCEDER-MENU.exe"])

    time.sleep(0.5)

    
#------------------------------------EXTRAER NOTAS DE ENTREGA DEL SMARTTOOLS------------------------------------
def exportar_notas_entrega(numero_orden):

    check_pause()
    exe = r"tiny_task\\exe"
    ruta_tiny_task = os.path.join(base_dir, exe)

    subprocess.run([f"{ruta_tiny_task}\\2-NE-IR-BUSCADOR.exe"])

    pyautogui.write(numero_orden, interval=0.1)
    presionar_teclas("enter")
    
    subprocess.run([f"{ruta_tiny_task}\\3-NE-EXPORTAR.exe"])

    pyautogui.press("backspace", presses=20)
    pyautogui.write(numero_orden, interval=0.1)
    presionar_teclas("enter")

    time.sleep(3)

    #el archivo .xls se guarda en la carpeta del a2 demo, lo vamos a mover a la carpeta del bot

    origen = f"C:\\a2DEMO\\{numero_orden}.xls"
    dest = f"Nota_Entrega\\{numero_orden}.xls"
    destino = os.path.join(base_dir, dest)

    shutil.move(origen, destino)

    combinar_excels(numero_orden)

    print(f"Nota de entrega: {numero_orden} exportada exitosamente.")






#-----------------------------------------------EXTRAER NOTAS DE ENTREGA-----------------------------
#Esta funcion es una de las principales, que llama a las demas funciones. esta funcion se conecta al deposito y extrae las notas de entrega
def notas_entrega():

    #Nos conectamos al deposito
    conectar_anydesk(direcciones_anydesk["DEPOSITO"], anydesk_password)
    check_pause()

    extraer_data() #Copias la data del anydesk y la pegas en local

    acceder_extraer_nota_entrega() #Aqui levantamos el smarttools

    for id_nota in libro_notas_entrega:
        exportar_notas_entrega(id_nota) #Bucle para extraer las notas de entrega

    envia_correo()

def ingresar_excel():

    carpeta_excel = os.path.join(base_dir, "excel_cargar")

    for archivo in os.listdir(carpeta_excel):
        if archivo.lower().endswith(".xlsx"):

            ruta_archivo = os.path.join(carpeta_excel, archivo)
            nombre_archivo = archivo.upper()

            for tienda, anydesk in direcciones_anydesk.items():

                if tienda in nombre_archivo:

                    #si estamos aqui, significa que tenemos el anydesk de la tienda detectada por el nombre
                    check_pause()
                    conectar_anydesk(anydesk, anydesk_password)
                    extraer_data()

                    excel_smarttools(ruta_archivo, nombre_archivo) #aqui convertimos el excel al formato smarttools







#--------------------------ESTA ES LA FUNCION MAESTRA QUE CONTROLARA TODAS LAS DEMAS--------------------------------------------

#PARA PAUSAR
threading.Thread(target=escuchar_pausa, daemon=True).start()

def main():

    controlador = int(input("Presione:\n1. Para trabajar con el correo.\n2.Para Insertar un archivo masivo:\n"))

    notas_entrega_pendiente = 0

    if(controlador == 1):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        #En la carpeta Data es en donde guardaremos los pdfs
        output_folder = os.path.join(base_dir, "pdf")
        os.makedirs(output_folder, exist_ok=True)

        output_folder_xlsx = os.path.join(base_dir, "excel_cargar")
        os.makedirs(output_folder, exist_ok=True)

        #Seleccionar la bandeja de entrada
        inbox = outlook.GetDefaultFolder(6) # 6 representa la bandeja de entrada

        folder_note_entrega = "Nota_Entrega"
        processed_folder_nota_entrega = inbox.Folders(folder_note_entrega)

        folder_carga_mercancia = "Completados"
        processed_folder_cargos = inbox.Folders(folder_carga_mercancia)

        messages = inbox.Items
        for i in range(messages.Count, 0, -1):
            message = messages.Item(i)
            try:
                subject = message.Subject #Titulo del correo
                sender = message.SenderName #Remitente

                print(f"Título: {subject} | Remitente: {sender}")

                #------------------------------------------ESTO ES PARA CUANDO ES NOTA DE ENTREGA---------------

                if fnmatch.fnmatch(subject, "*ENTREGA*"):

                    #Descargar y renombrar archivos adjuntos:
                    if message.Attachments.Count > 0:
                        for attachment in message.Attachments:

                            #Crear un nombre de archivo basado en el asunto del correo
                            filename = f"{subject} - {attachment.FileName}".replace(":", "-").replace("/", "-")
                            #filename = attachment.FileName.replace(":", "-").replace("/", "-")
                            #filename = f"{attachment.FileName} - {subject}".replace(":", "-").replace("/", "-")
                            file_path = os.path.join(output_folder, filename)

                            # Evitar sobrescritura
                            counter = 1
                            original_path = file_path
                            while os.path.exists(file_path):
                                base, ext = os.path.splitext(original_path)
                                file_path = f"{base} ({counter}){ext}"
                                counter += 1

                            # Guardar el archivo adjunto
                            attachment.SaveAsFile(file_path)
                            print(f"Archivo descargado: {file_path}")


                            #AQUI LLAMAMOS A LA FUNCION PARA EXTRAER EL NUMERO DE FACTURA
                            nombre_tienda, id_transaccion = pdf_reader(file_path)

                            #INSERTAMOS EN EL DICCIONARIO DE DATOS
                            libro_notas_entrega[id_transaccion] = nombre_tienda

                    # Mover el correo procesado a la carpeta "Completados"
                    message.Move(processed_folder_nota_entrega)
                    print(f"Correo movido a la carpeta '{folder_note_entrega}'")

                    destino = os.path.join(base_dir, "pdf", "completados")
                    shutil.move(file_path, destino)
                    print("Libro de notas de entrega:")
                    print(libro_notas_entrega)
                    notas_entrega_pendiente = 1


                #-------------ESTO ES PARA CUANDO ES UN EXCEL Y HAY QUE INGRESARSE---------------------------- 
                elif(fnmatch.fnmatch(subject, "*CARGAR*")):

                    #Descargar y renombrar archivos adjuntos:
                    if message.Attachments.Count > 0:
                        for attachment in message.Attachments:

                            #Crear un nombre de archivo basado en el asunto del correo
                            filename = f"{subject} - {attachment.FileName}".replace(":", "-").replace("/", "-")
                            #filename = attachment.FileName.replace(":", "-").replace("/", "-")
                            #filename = f"{attachment.FileName} - {subject}".replace(":", "-").replace("/", "-")
                            file_path = os.path.join(output_folder_xlsx, filename)

                            # Evitar sobrescritura
                            counter = 1
                            original_path = file_path
                            while os.path.exists(file_path):
                                base, ext = os.path.splitext(original_path)
                                file_path = f"{base} ({counter}){ext}"
                                counter += 1

                            # Guardar el archivo adjunto
                            attachment.SaveAsFile(file_path)
                            print(f"Archivo descargado: {file_path}")


                    # Mover el correo procesado a la carpeta "Completados"
                    message.Move(processed_folder_cargos) #RECUERDA CAMBIAR LOS NOMBRES DE LAS CARPETAS
                    print(f"Correo movido a la carpeta '{folder_carga_mercancia}'")

                    ingresar_excel()





                else:
                    print("No se encontraron NOTAS DE ENTREGA y/o COMPRAS DE MERCANCIA.")

            except Exception as e:
                print(f"Error procesando un correo: {e}")
                traceback.print_exc()

        if(notas_entrega_pendiente == 1):
            notas_entrega()
            print("fin")

    elif(controlador == 2):

        ruta_archivo = os.path.join(base_dir, "masivo")

        for archivo in os.listdir(ruta_archivo):

            for marca in lista_marcas:

                if fnmatch.fnmatch(archivo, f"*{marca}*"):
                    marca_actual = marca
                    print(f"Marca identificada: {marca_actual}")
                    break

            # conectar_anydesk(direcciones_anydesk["DEPOSITO"], anydesk_password)

            # time.sleep(5)

            # extraer_data()

            # direccion_archivo_completa = os.path.join(ruta_archivo, archivo)

            # cargar_smarttools_costo(archivo)


            for tienda in direcciones_anydesk:

                #aqui evaluamos el nombre con la marca para insertar
                if fnmatch.fnmatch(tienda, f"*{marca_actual}*"): # el nombre de la tienda coincide con la marca

                    conectar_anydesk(direcciones_anydesk[tienda], anydesk_password)
                    check_pause()
                    time.sleep(5)

                    extraer_data()
                    
                    direccion_archivo_completa = os.path.join(ruta_archivo, archivo)
                    check_pause()
                    cargar_smarttools_costo(archivo)

main()
