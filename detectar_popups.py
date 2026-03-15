import pygetwindow as gw
import time

def monitor_ventanas():
    print("=====================================================")
    print("  MONITOR DE VENTANAS EMERGENTES (POPUPS) DETECTADO")
    print("=====================================================")
    print("Cambiando de ventana... (Tienes 5 segundos para prepararte)")
    time.sleep(5)
    
    print("\n[*] Monitor iniciado. Registrando cualquier cambio en las ventanas...")
    print("Presiona Ctrl + C en esta consola para detener el script.\n")
    
    # Obtener la ventana actual
    active_window = gw.getActiveWindow()
    current_title = active_window.title if active_window else ""
    
    try:
        while True:
            # Revisar cuál es la ventana que está en primer plano ahora
            new_window = gw.getActiveWindow()
            if new_window:
                new_title = new_window.title
                
                # Si el título cambió, probablemente se abrió un popup o mensaje
                if new_title != current_title:
                    print(f"[!] ---> ¡DETECTÉ UNA NUEVA VENTANA!: '{new_title}'")
                    current_title = new_title
                    
            time.sleep(0.5)
            
    except KeyboardInterrupt:
        print("\nMonitor detenido por el usuario.")

if __name__ == "__main__":
    monitor_ventanas()
