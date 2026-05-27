import pdfplumber
import re

with pdfplumber.open("NOTA DE ENTREGA PURO EGO TOLON 03 04 26 - NOTA DE ENTREGA PURO EGO TOLON 03 04 26.pdf") as pdf:
    pagina = pdf.pages[0] # 0 = 1era Hoja
    texto = pagina.extract_text()
    

match = re.search(r"Cliente:\s*(.*)", texto)

match_id = re.search(r"CENTRO EMPRESARIAL STURGIS\S*(.*)", texto)

if match:
    cliente = match.group(1).strip()
    print(cliente)

if match_id:
    transaccion = match_id.group(1).strip()
    print(transaccion)


lineas = texto.split("\n")

numero_transaccion = None

for i, linea in enumerate(lineas):
    if "NOTA DE ENTREGA" in linea.upper():
        if i + 1 < len(lineas):
            posible = lineas[i+1].strip()

            if re.fullmatch(r"\d{8}", posible):
                numero_transaccion = posible
                break

#print(numero_transaccion)
#print(texto)