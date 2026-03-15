import tkinter as tk
from pynput import mouse
import time
import os

class CoordinateListTool:
    def __init__(self, target_width, target_height):
        self.root = tk.Tk()
        self.root.title(f"Coordenadas para {target_width}x{target_height}")
        self.root.attributes('-topmost', True)
        self.root.geometry("350x400")
        
        # Variables configuration
        self.target_width = target_width
        self.target_height = target_height
        
        # UI Elements
        self.label_inst = tk.Label(
            self.root, 
            text="1. Maximiza tu ventana (1600x900).\n2. CLIC DERECHO para guardar Clic normal en X,Y.\n3. CLIC CENTRAL (Rueda) para Leer y Escribir.\n¡Cuidado con salirte de los límites!", 
            font=("Arial", 10), justify=tk.LEFT
        )
        self.label_inst.pack(pady=10)
        
        self.label_curr = tk.Label(self.root, text="Actual: 0, 0", font=("Arial", 12, "bold"), fg="blue")
        self.label_curr.pack()
        
        self.label_warning = tk.Label(self.root, text="", fg="red", font=("Arial", 9, "bold"))
        self.label_warning.pack()
        
        # Scrollable Listbox for coordinates
        self.frame_list = tk.Frame(self.root)
        self.frame_list.pack(pady=10, fill=tk.BOTH, expand=True)
        
        self.scrollbar = tk.Scrollbar(self.frame_list)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(self.frame_list, font=("Consolas", 11), yscrollcommand=self.scrollbar.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        self.scrollbar.config(command=self.listbox.yview)

        # Buttons
        self.btn_clear = tk.Button(self.root, text="Borrar Lista", command=self.clear_list)
        self.btn_clear.pack(side=tk.LEFT, padx=20, pady=10)
        
        self.btn_auto = tk.Button(self.root, text="Generar Plantilla de Clics", command=self.generate_code, bg="lightblue")
        self.btn_auto.pack(side=tk.RIGHT, padx=20, pady=10)
        
        # State
        self.current_x = 0
        self.current_y = 0
        self.saved_coords = []
        
        # Start mouse listener
        self.listener = mouse.Listener(
            on_move=self.on_move,
            on_click=self.on_click)
        self.listener.start()
        
        self.update_ui()

    def on_move(self, x, y):
        self.current_x = int(x)
        self.current_y = int(y)

    def on_click(self, x, y, button, pressed):
        if pressed:
            cx, cy = int(x), int(y)
            # Solo guarda si está dentro del monitor de 1600x900
            if 0 <= cx <= self.target_width and 0 <= cy <= self.target_height:
                if button == mouse.Button.right:
                    action = "click"
                    desc = f"Clic Normal (x={cx}, y={cy})"
                elif button == mouse.Button.middle:
                    action = "read_write"
                    desc = f"Leer y Escribir (x={cx}, y={cy})"
                else:
                    return # No hacer nada si no es clic derecho o central
                
                coord_data = (action, cx, cy)
                self.saved_coords.append(coord_data)
                
                # Update Listbox immediately
                index = len(self.saved_coords)
                self.listbox.insert(tk.END, f"Paso {index}: {desc}")
                self.listbox.yview(tk.END) # Auto-scroll down
                
                # Guardar en log (append)
                with open("coordenadas_1600x900.txt", "a") as f:
                    f.write(f" Paso {index} [{action.upper()}] -> {cx}, {cy}\n")

    def update_ui(self):
        self.label_curr.config(text=f"Actual: X={self.current_x}, Y={self.current_y}")
        
        # Alert if mouse is outside the 1600x900 bounds
        if self.current_x > self.target_width or self.current_y > self.target_height or self.current_x < 0 or self.current_y < 0:
            self.label_warning.config(text=f"¡PELIGRO! Fuera del límite {self.target_width}x{self.target_height}")
        else:
            self.label_warning.config(text="")
            
        self.root.after(50, self.update_ui)
        
    def clear_list(self):
        self.saved_coords.clear()
        self.listbox.delete(0, tk.END)
        # Limpiar log
        if os.path.exists("coordenadas_1600x900.txt"):
            os.remove("coordenadas_1600x900.txt")
            
    def generate_code(self):
        """Genera un archivo con los clics en formato código Python"""
        if not self.saved_coords:
            return
            
        with open("codigo_generado.py", "w") as f:
            f.write("import pyautogui\n")
            f.write("import time\n\n")
            f.write("def ejecutar_clics(dato_a_escribir='DATO_DE_PRUEBA'):\n")
            for i, (action, x, y) in enumerate(self.saved_coords, 1):
                f.write(f"    # Paso {i}\n")
                if action == "click":
                    f.write(f"    pyautogui.click(x={x}, y={y})\n")
                elif action == "read_write":
                    f.write(f"    pyautogui.click(x={x}, y={y}) # Entrar al campo\n")
                    f.write(f"    time.sleep(0.5)\n")
                    f.write(f"    pyautogui.write(str(dato_a_escribir)) # Escribir el dato\n")
                f.write("    time.sleep(1) # Pausa por seguridad\n\n")
            f.write("if __name__ == '__main__':\n")
            f.write("    print('Ejecutando secuencia...')\n")
            f.write("    time.sleep(3)\n")
            f.write("    ejecutar_clics()\n")
        
        self.label_warning.config(text="¡Archivo codigo_generado.py creado!", fg="green")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = CoordinateListTool(target_width=1600, target_height=900)
    app.run()
