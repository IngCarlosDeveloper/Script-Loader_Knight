SmartTools Automation Bot
English
Overview

SmartTools Automation Bot is a Python-based automation system designed to streamline inventory, merchandise loading, delivery notes extraction, and synchronization workflows between multiple stores and a central warehouse using AnyDesk, Outlook, Excel, SmartTools, and A2Softway environments.

The project automates repetitive operational tasks such as:

Downloading PDF delivery notes from Outlook
Extracting transaction and store information from PDFs
Connecting remotely to stores and warehouse computers through AnyDesk
Compressing and transferring A2Softway databases
Importing and exporting SmartTools Excel files
Generating formatted Excel reports automatically
Sending processed files by email
Synchronizing merchandise and inventory information between systems

The goal of the project is to reduce manual work, minimize human error, and improve operational efficiency in retail inventory workflows.

Features
Email Automation
Reads Outlook inbox automatically
Detects emails related to:
Delivery Notes
Merchandise Loading
Downloads attachments automatically
Moves processed emails into organized folders
PDF Processing
Extracts:
Store names
Transaction IDs
Uses pdfplumber and regular expressions
Remote Synchronization
Connects remotely using AnyDesk
Transfers compressed database files
Executes remote batch scripts automatically
SmartTools Integration
Opens SmartTools automatically
Imports and exports inventory files
Loads merchandise with or without costs
Automates UI interaction using:
PyAutoGUI
PyDirectInput
TinyTask macros
Excel Automation
Converts XLSX to XLS
Applies formatting automatically
Generates inventory-ready spreadsheets
Uses:
Pandas
OpenPyXL
Win32 COM automation
Process Control
Pause/resume execution using keyboard shortcuts
Multithreaded execution control
Technologies Used
Python
PyAutoGUI
PyDirectInput
Pandas
OpenPyXL
pdfplumber
Win32COM
Requests
AnyDesk
SmartTools
Outlook COM API
Project Structure
project/
│
├── excel_cargar/
├── Nota_Entrega/
├── plantilla/
├── pdf/
├── Data/
├── tiny_task/
├── imagenes_pyautogui/
│
├── main.py
├── README.md
Main Functionalities
Delivery Notes Workflow
Read incoming emails
Download attached PDFs
Extract store and transaction information
Connect to warehouse computer
Synchronize A2Softway database
Export SmartTools delivery notes
Generate formatted Excel files
Send reports back to stores
Merchandise Loading Workflow
Detect merchandise loading emails
Download Excel attachments
Connect remotely to target store
Synchronize database
Convert and prepare SmartTools-compatible files
Load inventory into SmartTools
Synchronize updated data back to remote machine
Requirements
Python Version
Python 3.10+
Required Packages

Install dependencies:

pip install pyautogui
pip install pydirectinput
pip install pandas
pip install openpyxl
pip install pdfplumber
pip install pyperclip
pip install pywin32
pip install requests
pip install xlrd
Windows Requirements

This project is designed specifically for Windows environments because it depends on:

Outlook COM API
Win32 Clipboard APIs
SmartTools
AnyDesk
Batch scripts
Windows Explorer automation
Important Notes
Screen resolutions and UI coordinates may need adjustment depending on the machine.
TinyTask macros must match the SmartTools UI exactly.
AnyDesk IDs and passwords should be configured securely.
This project automates critical operational workflows and should be tested carefully before production use.
Security Recommendations

Do not expose:

AnyDesk passwords
Store credentials
Outlook credentials
Internal business information

Use environment variables or encrypted configuration files for sensitive data.

Future Improvements

Possible future improvements include:

Replacing screen automation with direct APIs
Adding logging system
Database integration
Web dashboard
Error recovery system
Multi-user support
AI-assisted document validation
Author

Carlos Lopez

Python Developer | Automation Developer | Robotics Enthusiast

Español
Descripción General

SmartTools Automation Bot es un sistema de automatización desarrollado en Python diseñado para optimizar flujos de trabajo relacionados con inventario, carga de mercancía, extracción de notas de entrega y sincronización de información entre tiendas y almacenes utilizando AnyDesk, Outlook, Excel, SmartTools y entornos A2Softway.

El proyecto automatiza tareas operativas repetitivas como:

Descargar notas de entrega PDF desde Outlook
Extraer información de tiendas y transacciones
Conectarse remotamente mediante AnyDesk
Comprimir y transferir bases de datos A2Softway
Importar y exportar archivos Excel desde SmartTools
Generar reportes automáticamente
Enviar archivos procesados por correo
Sincronizar mercancía e inventario entre sistemas

El objetivo principal es reducir trabajo manual, minimizar errores humanos y mejorar la eficiencia operativa.

Características
Automatización de Correos
Lectura automática de Outlook
Detección de correos relacionados con:
Notas de entrega
Carga de mercancía
Descarga automática de archivos adjuntos
Organización automática de correos procesados
Procesamiento de PDFs
Extracción de:
Nombre de tienda
ID de transacción
Uso de pdfplumber y expresiones regulares
Sincronización Remota
Conexión remota mediante AnyDesk
Transferencia de bases de datos comprimidas
Ejecución automática de scripts .bat
Integración con SmartTools
Apertura automática de SmartTools
Importación y exportación de inventario
Carga de mercancía con y sin costo
Automatización de interfaz gráfica usando:
PyAutoGUI
PyDirectInput
Macros TinyTask
Automatización Excel
Conversión de XLSX a XLS
Aplicación automática de formatos
Generación de hojas listas para inventario
Uso de:
Pandas
OpenPyXL
Automatización COM de Excel
Control del Proceso
Pausa y reanudación mediante teclado
Control multihilo
Tecnologías Utilizadas
Python
PyAutoGUI
PyDirectInput
Pandas
OpenPyXL
pdfplumber
Win32COM
Requests
AnyDesk
SmartTools
Outlook COM API
Estructura del Proyecto
project/
│
├── excel_cargar/
├── Nota_Entrega/
├── plantilla/
├── pdf/
├── Data/
├── tiny_task/
├── imagenes_pyautogui/
│
├── main.py
├── README.md
Funcionalidades Principales
Flujo de Notas de Entrega
Leer correos entrantes
Descargar PDFs adjuntos
Extraer información de tienda y transacción
Conectarse al almacén remoto
Sincronizar base de datos A2Softway
Exportar notas de entrega desde SmartTools
Generar archivos Excel formateados
Enviar reportes nuevamente a las tiendas
Flujo de Carga de Mercancía
Detectar correos de carga de mercancía
Descargar archivos Excel
Conectarse remotamente a la tienda
Sincronizar base de datos
Preparar archivos compatibles con SmartTools
Cargar inventario automáticamente
Sincronizar información actualizada
Requisitos
Versión de Python
Python 3.10+
Dependencias

Instalar paquetes:

pip install pyautogui
pip install pydirectinput
pip install pandas
pip install openpyxl
pip install pdfplumber
pip install pyperclip
pip install pywin32
pip install requests
pip install xlrd
Requisitos de Windows

Este proyecto está diseñado específicamente para Windows porque depende de:

Outlook COM API
APIs Win32
SmartTools
AnyDesk
Scripts Batch
Automatización del Explorador de Windows
Notas Importantes
Las coordenadas de pantalla pueden requerir ajustes según la resolución.
Las macros TinyTask deben coincidir exactamente con la interfaz de SmartTools.
Las credenciales AnyDesk deben almacenarse de forma segura.
El sistema automatiza procesos críticos y debe probarse cuidadosamente antes de producción.
Recomendaciones de Seguridad

No exponer:

Contraseñas de AnyDesk
Credenciales internas
Información empresarial sensible
Credenciales de Outlook

Se recomienda usar variables de entorno o archivos cifrados para información sensible.

Mejoras Futuras

Posibles mejoras:

Reemplazar automatización visual por APIs directas
Sistema de logs
Integración con base de datos
Dashboard web
Recuperación automática de errores
Soporte multiusuario
Validación documental con IA
Autor

Carlos Lopez

Desarrollador Python | Desarrollador de Automatización | Entusiasta de la Robótica
