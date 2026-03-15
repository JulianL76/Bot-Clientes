import pyautogui
import pandas as pd
import time
import sys

# ==========================================
# CONFIGURACIÓN INICIAL
# ==========================================
# Ruta de tu archivo Excel
EXCEL_FILE = 'datos.xlsx'

# Nombre de la hoja de cálculo (dejar en None si es la primera hoja)
SHEET_NAME = 0 

# Pausa estándar entre cada acción del bot (en segundos)
pyautogui.PAUSE = 1.0

# Medida de seguridad: Si mueves el mouse a una de las esquinas de la pantalla, el bot se detendrá
pyautogui.FAILSAFE = True

# ==========================================
# FUNCIONES DEL BOT
# ==========================================

def read_data_from_excel(file_path):
    """
    Lee los datos desde un archivo Excel y retorna un DataFrame de Pandas.
    """
    try:
        print(f"[*] Leyendo archivo Excel: {file_path}...")
        df = pd.read_excel(file_path, sheet_name=SHEET_NAME)
        print(f"[+] Se cargaron {len(df)} filas correctamente.")
        return df
    except Exception as e:
        print(f"[-] Error al leer el archivo Excel: {e}")
        sys.exit(1)

def process_row(row):
    """
    Contiene la lógica de navegación y clics para UNA fila de datos del Excel.
    Reemplaza las coordenadas y acciones con las que necesites.
    """
    print(f"[*] Procesando fila: {row.to_dict()}")
    
    # -------------------------------------------------------------
    # EXTRACCIÓN DE DATOS DE EXCEL (Reemplaza los nombres de columnas)
    # -------------------------------------------------------------
    
    # Ejemplo:
    dato_paso_2 = str(row.get('Columna_Paso_2', 'Vacio'))
    dato_paso_4 = str(row.get('Columna_Paso_4', 'Vacio'))
    dato_paso_6 = str(row.get('Columna_Paso_6', 'Vacio'))
    dato_paso_8 = str(row.get('Columna_Paso_8', 'Vacio'))
    dato_paso_10 = str(row.get('Columna_Paso_10', 'Vacio'))
    dato_paso_12 = str(row.get('Columna_Paso_12', 'Vacio'))
    dato_paso_14 = str(row.get('Columna_Paso_14', 'Vacio'))

    # -------------------------------------------------------------
    # SECUENCIA DE LLENADO DE FORMULARIO
    # -------------------------------------------------------------
    
    # Paso 1 [CLICK]
    pyautogui.click(x=386, y=117)
    time.sleep(0.5)

    # Paso 2 [READ_WRITE]
    pyautogui.click(x=386, y=117)
    time.sleep(0.5)
    pyautogui.write(dato_paso_2)
    time.sleep(0.5)

    # Paso 3 [CLICK]
    pyautogui.click(x=392, y=143)
    time.sleep(0.5)

    # Paso 4 [READ_WRITE]
    pyautogui.click(x=392, y=143)
    time.sleep(0.5)
    pyautogui.write(dato_paso_4)
    time.sleep(0.5)

    # Paso 5 [CLICK]
    pyautogui.click(x=383, y=170)
    time.sleep(0.5)

    # Paso 6 [READ_WRITE]
    pyautogui.click(x=383, y=170)
    time.sleep(0.5)
    pyautogui.write(dato_paso_6)
    time.sleep(0.5)

    # Paso 7 [CLICK]
    pyautogui.click(x=979, y=171)
    time.sleep(0.5)

    # Paso 8 [READ_WRITE]
    pyautogui.click(x=965, y=172)
    time.sleep(0.5)
    pyautogui.write(dato_paso_8)
    time.sleep(0.5)

    # Paso 9 [CLICK]
    pyautogui.click(x=792, y=198)
    time.sleep(0.5)

    # Paso 10 [READ_WRITE]
    pyautogui.click(x=792, y=198)
    time.sleep(0.5)
    pyautogui.write(dato_paso_10)
    time.sleep(0.5)

    # Paso 11 [CLICK]
    pyautogui.click(x=421, y=224)
    time.sleep(0.5)

    # Paso 12 [READ_WRITE]
    pyautogui.click(x=564, y=227)
    time.sleep(0.5)
    pyautogui.write(dato_paso_12)
    time.sleep(0.5)

    # Paso 13 [CLICK]
    pyautogui.click(x=780, y=228)
    time.sleep(0.5)

    # Paso 14 [READ_WRITE]
    pyautogui.click(x=780, y=228)
    time.sleep(0.5)
    pyautogui.write(dato_paso_14)
    time.sleep(0.5)
    
    # Pequeña pausa opcional si la página o sistema tarda en cargar
    time.sleep(1)
    
    print("[+] Fila procesada con éxito.\n")

def open_form_sequence():
    """
    Secuencia de clics iniciales para abrir el formulario antes de llenar los datos.
    """
    print("[*] Ejecutando secuencia inicial para abrir el formulario...")
    
    # Paso 1 -> 70, 176
    print("    - Clic 1 (70, 176)")
    pyautogui.click(x=70, y=176)
    time.sleep(1) # Pausa para que la interfaz reaccione (ajústalo si es necesario)
    
    # Paso 2 -> 52, 205
    print("    - Clic 2 (52, 205)")
    pyautogui.click(x=52, y=205)
    time.sleep(1)
    
    # Paso 3 -> 281, 85
    print("    - Clic 3 (281, 85)")
    pyautogui.click(x=281, y=85)
    time.sleep(2) # Pausa más larga para esperar que cargue el formulario completo
    
    print("[+] Formulario abierto y listo.\n")


def main():
    print("===========================================")
    print("    INICIANDO BOT DE AUTOMATIZACIÓN")
    print("===========================================")
    print("Tienes 5 segundos para prepararte y cambiar a la ventana deseada...")
    time.sleep(5)
    
    # 0. Ejecutar secuencia inicial para abrir el formulario
    open_form_sequence()
    
    # 1. Leer datos del Excel
    df = read_data_from_excel(EXCEL_FILE)
    
    # 2. Iterar sobre cada fila del Excel
    for index, row in df.iterrows():
        print(f"--- Iniciando iteración {index + 1} de {len(df)} ---")
        try:
            process_row(row)
        except Exception as e:
            print(f"[-] Error procesando la fila {index + 1}: {e}")
            print("[!] Continuando con la siguiente fila...\n")
            continue
            
    print("===========================================")
    print("    BOT FINALIZADO CON ÉXITO")
    print("===========================================")

if __name__ == "__main__":
    main()
