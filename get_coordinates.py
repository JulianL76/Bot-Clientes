import pyautogui
import time
import sys

print("==================================================")
print("   HERRAMIENTA PARA OBTENER COORDENADAS (X, Y)    ")
print("==================================================")
print("Mueve el mouse por la pantalla para ver las coordenadas.")
print("Presiona Ctrl + C en esta ventana para salir.")
print("--------------------------------------------------\n")

try:
    while True:
        # Obtiene las coordenadas actuales del mouse
        x, y = pyautogui.position()
        
        # Formatea el texto para que se actualice en la misma línea
        position_str = f"Coordenadas actuales ->  X: {str(x).rjust(4)}   Y: {str(y).rjust(4)}"
        
        # Imprime la posición y borra la línea para la siguiente iteración
        sys.stdout.write('\r' + position_str)
        sys.stdout.flush()
        
        # Pequeña pausa para no sobrecargar el procesador
        time.sleep(0.1)
        
except KeyboardInterrupt:
    print("\n\n[+] Programa terminado. ¡Espero que hayas encontrado tus coordenadas!")
