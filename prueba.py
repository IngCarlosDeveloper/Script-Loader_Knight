import threading
import time
import keyboard

pause_event = threading.Event()
pause_event.set()  # empieza en RUN

def toggle_pause():
    while True:
        keyboard.wait('|')  # tecla de pausa
        if pause_event.is_set():
            pause_event.clear()
            print("\n⏸ PAUSADO\n")
        else:
            pause_event.set()
            print("\n▶ REANUDADO\n")

def check_pause():
    pause_event.wait()

def proceso_prueba():
    i = 0
    while True:
        check_pause()
        print(f"Procesando... {i}")
        i += 1
        time.sleep(1)

# Thread que escucha la tecla
threading.Thread(target=toggle_pause, daemon=True).start()

# Proceso principal
proceso_prueba()