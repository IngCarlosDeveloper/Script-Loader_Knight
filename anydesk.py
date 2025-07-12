import os
import subprocess
import pyperclip
import fnmatch
import time
import shutil

# Ruta donde se encuentran los archivos Excel descargados
excel_folder = r"C:\Users\user\Documents\Trabajo\Mercancia"  # Cambia esta ruta

# Diccionario con las direcciones de AnyDesk según la tienda
anydesk_addresses = {
    "NAME_A":"ANYDESK_NUMBER",
    "NAME_B":"ANYDESK_NUMBER"
}

# Listar los archivos Excel en la carpeta
excel_files = [f for f in os.listdir(excel_folder) if f.endswith(".xlsx")]

# Procesar cada archivo Excel
for file in excel_files:
    try:
        # Extraer el nombre del archivo sin la extensión
        file_name = os.path.splitext(file)[0]  # Nombre sin extensión
        archivo = file_name
        print(f"Procesando archivo: {file_name}")

        # Buscar la tienda en el nombre del archivo usando un patrón con fnmatch
        for store_name in anydesk_addresses.keys():
            if fnmatch.fnmatch(file_name, f"*{store_name}*"):  # Usa * como comodín
                print(f"Tienda identificada: {store_name}")

                # Copiar la dirección de AnyDesk correspondiente al portapapeles
                anydesk_address = anydesk_addresses[store_name]
                pyperclip.copy(anydesk_address)
                print(f"Dirección de AnyDesk '{anydesk_address}' copiada al portapapeles.")
                break
        else:
            print(f"No se encontró una tienda válida en el archivo: {file_name}")
            break
    except Exception as e:
        print(f"Error procesando el archivo '{file}': {e}")
        break

    origen = r"C:\\Users\\user\\Documents\\Trabajo\\Mercancia\\"+file_name+".xlsx"

    destino = r"C:\\Users\\user\\Documents\\Trabajo\\Mercancia\\utilitarios\\datos.xlsx"

    desecho = r"C:\\Users\\user\\Documents\\Trabajo\\Mercancia\\Finalizados\\"+file_name+".xlsx"

    shutil.copy(origen, destino)

    try:
        shutil.move(origen, desecho)
        print(f"El archivo se ha movido a: {desecho}")
    except FileNotFoundError:
        print("El archivo no se encontró en la ubicación especificada.")
    except PermissionError:
        print("No tienes permisos para mover este archivo.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")

    subprocess.Popen(["C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe"])

    time.sleep(3)

    subprocess.run(["C:\\Users\\user\\Desktop\\Loader_Knight\\Tiny Task\\exe\\conecta_anydesk.exe"])

    time.sleep(10)

    subprocess.run(["C:\\Users\\user\\Desktop\\Loader_Knight\\Tiny Task\\exe\\cmd.exe"])

    time.sleep(3)

    subprocess.run(["C:\\Users\\user\\Desktop\\Loader_Knight\\Tiny Task\\exe\\delete.exe"])

    time.sleep(3)

    subprocess.run(["C:\\Users\\user\\Desktop\\Loader_Knight\\Tiny Task\\exe\\cmd.exe"])

    time.sleep(3)

    subprocess.run(["C:\\Users\\user\\Desktop\\Loader_Knight\\Tiny Task\\exe\\comprimidor.exe"])

    time.sleep(20)

    subprocess.run(["C:\\Users\\user\\Desktop\\Loader_Knight\\Tiny Task\\exe\\copiar.exe"])

    time.sleep(45)

    os.system('taskkill /F /im AnyDesk.exe"')

    subprocess.run([r"C:\\descomprimidor.bat"])

    time.sleep(10)

    os.system('taskkill /F /im explorer.exe"')
    time.sleep(0.5)
    os.system("start explorer.exe")

    time.sleep(1)
    
    os.chdir(r"C:\\a2Softway")

    subprocess.Popen(["a2Admin.exe"])

    time.sleep(1.5)
    subprocess.run(["C:\\Users\\user\\Desktop\\Loader_Knight\\Tiny Task\\exe\\login.exe"])

    time.sleep(40)

    os.system('taskkill /F /im a2Admin.exe"')
    time.sleep(2)

    subprocess.run(["python", "C:\\Users\\user\\Desktop\\Loader_Knight\\excel.py"])
