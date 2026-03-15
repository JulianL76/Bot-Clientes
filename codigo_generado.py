import pyautogui
import time

def ejecutar_clics(dato_a_escribir='DATO_DE_PRUEBA'):
    # Paso 1
    pyautogui.click(x=668, y=417)
    time.sleep(1) # Pausa por seguridad

    # Paso 2
    pyautogui.click(x=44, y=55)
    time.sleep(1) # Pausa por seguridad

if __name__ == '__main__':
    print('Ejecutando secuencia...')
    time.sleep(3)
    ejecutar_clics()
