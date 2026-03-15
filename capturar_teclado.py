"""
capturar_teclado.py
-------------------
Captura cada pulsación de tecla con timestamp exacto (milisegundos)
y la guarda en log_teclado_manual.txt, en el mismo formato que log_teclas.txt.

Úsalo así:
  1. Abre una consola y ejecuta:  python capturar_teclado.py
  2. Cambia a ABAKO y navega el formulario A MANO como lo harías normalmente
  3. Cuando termines, vuelve a la consola y presiona ESC para detener

Luego compara log_teclado_manual.txt con log_teclas.txt para ver la diferencia.
"""

import time
import sys

try:
    from pynput import keyboard
except ImportError:
    print("[-] Falta la librería 'pynput'. Instalando...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pynput"])
    from pynput import keyboard

LOG_FILE = 'log_teclado_manual.txt'
log = None
capturando = True
ultimo_ts = None  # Para calcular el tiempo entre teclas

def ts_ahora():
    """Devuelve timestamp con milisegundos."""
    return time.strftime('%H:%M:%S') + f'.{int(time.time() * 1000) % 1000:03d}'

def tiempo_desde_ultima(ts_actual):
    """Calcula cuántos ms pasaron desde la última tecla."""
    global ultimo_ts
    if ultimo_ts is None:
        ultimo_ts = time.time()
        return ""
    delta_ms = (time.time() - ultimo_ts) * 1000
    ultimo_ts = time.time()
    if delta_ms > 300:  # Solo muestra el gap si es mayor a 300ms (relevante)
        return f"  (+{delta_ms:.0f}ms)"
    return ""

def nombre_tecla(key):
    """Convierte la tecla a un nombre legible similar al formato del bot."""
    try:
        # Tecla de carácter normal
        c = key.char
        if c is None:
            return repr(key)
        return repr(c)
    except AttributeError:
        # Tecla especial (enter, tab, shift, etc.)
        nombre = str(key).replace('Key.', '')
        return f"'{nombre}'"

def on_press(key):
    global capturando
    
    # ESC detiene la captura
    if key == keyboard.Key.esc:
        capturando = False
        ts = ts_ahora()
        log.write(f'[{ts}] --- CAPTURA DETENIDA (ESC) ---\n')
        log.flush()
        return False  # Para el listener
    
    ts = ts_ahora()
    gap = tiempo_desde_ultima(ts)
    nombre = nombre_tecla(key)
    
    linea = f'[{ts}] PRESS    {nombre}{gap}\n'
    log.write(linea)
    log.flush()
    
    # También mostrar en consola para feedback visual
    print(f'  {linea.strip()}')

def on_release(key):
    pass  # No logueamos releases para mantener el mismo formato que el bot

def main():
    global log
    
    print("=" * 60)
    print("   CAPTURADOR DE TECLADO MANUAL")
    print("=" * 60)
    print(f"[*] Guardando en: {LOG_FILE}")
    print("[*] Formato idéntico a log_teclas.txt del bot")
    print()
    print("  INSTRUCCIONES:")
    print("  1. Este script empieza a capturar INMEDIATAMENTE")
    print("  2. Cambia a ABAKO y navega el formulario a mano")
    print("  3. Cuando termines, presiona ESC para detener")
    print()
    print("  Capturando... (puedes cambiar de ventana ahora)")
    print("-" * 60)
    
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        log = f
        log.write('\n' + '=' * 60 + '\n')
        log.write(f'  TECLADO MANUAL: {time.strftime("%Y-%m-%d %H:%M:%S")}\n')
        log.write('  (Los gaps +Xms entre líneas muestran cuánto tardaste entre teclas)\n')
        log.write('=' * 60 + '\n')
        log.flush()
        
        # Iniciar listener de teclado (captura global, funciona en cualquier ventana)
        with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
            listener.join()
    
    print()
    print("-" * 60)
    print(f"[+] Captura finalizada. Revisa: {LOG_FILE}")
    print()
    print("  CÓMO COMPARAR:")
    print("  - log_teclado_manual.txt = lo que TÚ haces")
    print("  - log_teclas.txt         = lo que el BOT hace")
    print("  Busca diferencias en el orden de teclas y los gaps de tiempo.")

if __name__ == "__main__":
    main()
