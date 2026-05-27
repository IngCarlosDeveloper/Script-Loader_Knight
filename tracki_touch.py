from pynput import mouse

def imprimir_clicks():

    def on_click(x, y, button, pressed):

        if pressed:
            print(f"Click en: X={x} Y={y}")


    with mouse.Listener(on_click=on_click) as listener:
        listener.join()


imprimir_clicks()