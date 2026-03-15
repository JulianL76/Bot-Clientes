import tkinter as tk
from pynput import mouse
import time

class CoordinateTool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Coordenadas Multi-Pantalla")
        
        # Hacer la ventana siempre visible (Always on top)
        self.root.attributes('-topmost', True)
        self.root.geometry("400x150")
        
        self.label_inst = tk.Label(self.root, text="Mueve el mouse para ver coordenadas.\n¡Haz CLIC DERECHO para GUARDAR una posición!", font=("Arial", 10))
        self.label_inst.pack(pady=10)
        
        self.label_coord = tk.Label(self.root, text="X: 0 | Y: 0", font=("Arial", 16, "bold"), fg="blue")
        self.label_coord.pack(pady=5)
        
        self.label_saved = tk.Label(self.root, text="Guardado:\n", font=("Arial", 9), fg="green")
        self.label_saved.pack(pady=5)
        
        # Variables de estado
        self.current_x = 0
        self.current_y = 0
        self.saved_coords = []
        
        # Iniciar listener del mouse de pynput 
        # (pynput maneja los monitores múltiples mucho mejor que pyautogui para leer eventos puros del SO)
        self.listener = mouse.Listener(
            on_move=self.on_move,
            on_click=self.on_click)
        self.listener.start()
        
        # Bucle de actualización de la UI
        self.update_ui()

    def on_move(self, x, y):
        self.current_x = int(x)
        self.current_y = int(y)

    def on_click(self, x, y, button, pressed):
        if pressed and button == mouse.Button.right:
            coord = f"X={int(x)}, Y={int(y)}"
            self.saved_coords.append(coord)
            
            # Guardamos en un archivo de texto por si acaso
            with open("coordenadas_guardadas.txt", "a") as f:
                f.write(f"{coord} - {time.strftime('%H:%M:%S')}\n")

    def update_ui(self):
        # Actualizar etiqueta en tiempo real
        self.label_coord.config(text=f"X: {self.current_x} | Y: {self.current_y}")
        
        # Mostrar las últimas 3 guardadas
        if self.saved_coords:
            recent = "\\n".join(self.saved_coords[-3:])
            self.label_saved.config(text=f"Últimas guardadas:\\n{recent}")
            
        self.root.after(50, self.update_ui)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = CoordinateTool()
    app.run()
