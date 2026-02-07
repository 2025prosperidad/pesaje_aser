import tkinter as tk
from tkinter import ttk, scrolledtext
import serial
import serial.tools.list_ports
import threading
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image, ImageTk
import io
import time

class WeightMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("Monitor de Peso - Hik-Connect")
        self.root.geometry("1600x900")
        self.root.configure(bg="#1e1e1e")

        self.serial_port = None
        self.is_running = False
        self.current_weight = 0
        self.status = "ST"
        self.driver = None
        self.browser_frame = None

        self.setup_ui()
        self.log_message("=== INICIANDO APLICACIÓN ===")
        self.root.after(1000, self.init_selenium)

    def setup_ui(self):
        # Frame superior - Configuración
        config_frame = tk.Frame(self.root, bg="#2d2d2d", padx=10, pady=10)
        config_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(config_frame, text="Puerto:", bg="#2d2d2d", fg="white").grid(row=0, column=0, padx=5)
        self.port_combo = ttk.Combobox(config_frame, width=10, state="readonly")
        self.port_combo.grid(row=0, column=1, padx=5)
        self.refresh_ports()

        tk.Label(config_frame, text="Baud Rate:", bg="#2d2d2d", fg="white").grid(row=0, column=2, padx=5)
        self.baud_combo = ttk.Combobox(config_frame, width=10, values=["1200", "9600", "19200", "38400", "115200"], state="readonly")
        self.baud_combo.set("1200")
        self.baud_combo.grid(row=0, column=3, padx=5)

        self.btn_refresh = tk.Button(config_frame, text="🔄", command=self.refresh_ports, bg="#3d3d3d", fg="white")
        self.btn_refresh.grid(row=0, column=4, padx=2)

        self.btn_connect = tk.Button(config_frame, text="Conectar", command=self.toggle_connection,
                                     bg="#0d7377", fg="white", width=12, font=("Arial", 10, "bold"))
        self.btn_connect.grid(row=0, column=5, padx=10)

        # Frame principal horizontal: Izquierda (Peso) y Derecha (Navegador)
        main_horizontal_frame = tk.Frame(self.root, bg="#1e1e1e")
        main_horizontal_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # === LADO IZQUIERDO: Monitor de Peso ===
        left_frame = tk.Frame(main_horizontal_frame, bg="#1e1e1e", width=400)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))
        left_frame.pack_propagate(False)

        # Display grande del peso
        display_frame = tk.Frame(left_frame, bg="#1e1e1e")
        display_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Indicador de estado
        self.status_label = tk.Label(display_frame, text="●", font=("Arial", 40),
                                     bg="#1e1e1e", fg="#ff4444")
        self.status_label.pack(pady=5)

        self.status_text = tk.Label(display_frame, text="DESCONECTADO", font=("Arial", 14),
                                   bg="#1e1e1e", fg="#888888")
        self.status_text.pack()

        # Display del peso
        self.weight_display = tk.Label(display_frame, text="0", font=("Arial", 80, "bold"),
                                       bg="#1e1e1e", fg="#00ff00")
        self.weight_display.pack(pady=20)

        self.unit_label = tk.Label(display_frame, text="kg", font=("Arial", 30),
                                   bg="#1e1e1e", fg="#888888")
        self.unit_label.pack()

        # Información adicional
        info_frame = tk.Frame(display_frame, bg="#2d2d2d")
        info_frame.pack(fill=tk.X, padx=20, pady=10)

        self.stability_label = tk.Label(info_frame, text="Estado: --", font=("Arial", 12),
                                       bg="#2d2d2d", fg="white")
        self.stability_label.pack(side=tk.LEFT, padx=20)

        self.type_label = tk.Label(info_frame, text="Tipo: --", font=("Arial", 12),
                                   bg="#2d2d2d", fg="white")
        self.type_label.pack(side=tk.LEFT, padx=20)

        # Frame inferior - Log (en el lado izquierdo)
        log_frame = tk.LabelFrame(left_frame, text="Registro de Datos", bg="#2d2d2d",
                                 fg="white", font=("Arial", 10))
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, bg="#1a1a1a",
                                                  fg="#00ff00", font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # === LADO DERECHO: Navegador Embebido ===
        right_container = tk.Frame(main_horizontal_frame, bg="#2d2d2d")
        right_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Título
        title_frame = tk.Frame(right_container, bg="#2d2d2d", height=30)
        title_frame.pack(fill=tk.X, side=tk.TOP)
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text="Hik-Connect - Tiempo Real", 
                bg="#2d2d2d", fg="white", font=("Arial", 11, "bold")).pack(pady=5)

        # Área para mostrar el navegador embebido - USAR CANVAS
        self.browser_canvas = tk.Canvas(right_container, bg="black", highlightthickness=0)
        self.browser_canvas.pack(fill=tk.BOTH, expand=True)

        # Botones inferiores
        btn_frame = tk.Frame(self.root, bg="#1e1e1e")
        btn_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(btn_frame, text="Limpiar Log", command=self.clear_log,
                 bg="#3d3d3d", fg="white").pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Guardar Datos", command=self.save_log,
                 bg="#3d3d3d", fg="white").pack(side=tk.LEFT, padx=5)

    def refresh_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo['values'] = ports
        if ports:
            if 'COM5' in ports:
                self.port_combo.set('COM5')
            else:
                self.port_combo.set(ports[0])

    def init_selenium(self):
        try:
            self.log_message("Iniciando navegador embebido con Selenium...")
            chrome_options = Options()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)

            # Usar webdriver-manager para descargar automáticamente ChromeDriver
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
            self.driver.get("https://www.hik-connect.com/views/login/index.html#/portal")

            self.log_message("Navegador embebido iniciado correctamente.")
            self.update_browser_view()
        except Exception as e:
            self.log_message(f"Error al iniciar Selenium: {e}")
            self.browser_canvas.config(text=f"Error: {str(e)}", fg="red")

    def update_browser_view(self):
        try:
            if self.driver:
                # Forzar actualización de geometría
                self.root.update_idletasks()
                
                # Obtener dimensiones actuales del canvas
                canvas_width = self.browser_canvas.winfo_width()
                canvas_height = self.browser_canvas.winfo_height()
                
                # Si el canvas aún no tiene dimensiones válidas, esperar y reintentar
                if canvas_width <= 1 or canvas_height <= 1:
                    self.root.after(500, self.update_browser_view)
                    return
                
                # Capturar pantalla del navegador
                screenshot = self.driver.get_screenshot_as_png()
                image = Image.open(io.BytesIO(screenshot))
                
                # Redimensionar para llenar todo el espacio disponible
                image = image.resize((canvas_width, canvas_height), Image.LANCZOS)
                photo = ImageTk.PhotoImage(image)
                
                # Limpiar canvas y dibujar imagen
                self.browser_canvas.delete("all")
                self.browser_canvas.create_image(0, 0, image=photo, anchor=tk.NW)
                self.browser_canvas.image = photo  # Mantener referencia
                
                self.root.after(1000, self.update_browser_view)
        except Exception as e:
            self.log_message(f"Error al actualizar vista del navegador: {e}")

    def toggle_connection(self):
        if not self.is_running:
            self.connect()
        else:
            self.disconnect()

    def connect(self):
        try:
            port = self.port_combo.get()
            baud = int(self.baud_combo.get())

            self.serial_port = serial.Serial(
                port=port,
                baudrate=baud,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )

            self.is_running = True
            self.btn_connect.config(text="Desconectar", bg="#d32f2f")
            self.status_label.config(fg="#00ff00")
            self.status_text.config(text="CONECTADO", fg="#00ff00")

            # Iniciar thread de lectura
            self.read_thread = threading.Thread(target=self.read_serial, daemon=True)
            self.read_thread.start()

            self.log_message(f"Conectado a {port} @ {baud} baud")

        except Exception as e:
            self.log_message(f"Error al conectar: {str(e)}")

    def disconnect(self):
        self.is_running = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()

        self.btn_connect.config(text="Conectar", bg="#0d7377")
        self.status_label.config(fg="#ff4444")
        self.status_text.config(text="DESCONECTADO", fg="#888888")
        self.log_message("Desconectado")

    def read_serial(self):
        while self.is_running:
            try:
                if self.serial_port and self.serial_port.is_open:
                    line = self.serial_port.readline().decode('utf-8', errors='ignore').strip()
                    if line:
                        self.root.after(0, self.process_data, line)
            except Exception as e:
                self.log_message(f"Error de lectura: {str(e)}")
                break

    def process_data(self, data):
        self.log_message(f"Datos recibidos: {data}")
        match = re.match(r'([A-Z]{2}),([A-Z]{2}),([+\-])\s*(\d+)kg', data)
        if match:
            status, weight_type, sign, weight = match.groups()
            weight_value = int(weight)
            self.root.after(0, self.update_display, weight_value, status, weight_type)

    def update_display(self, weight, status, weight_type):
        self.current_weight = weight
        self.status = status
        self.weight_display.config(text=str(weight))
        if status == "ST":
            self.stability_label.config(text="Estado: ESTABLE", fg="#00ff00")
            self.weight_display.config(fg="#00ff00")
        else:
            self.stability_label.config(text="Estado: INESTABLE", fg="#ffaa00")
            self.weight_display.config(fg="#ffaa00")
        type_text = "BRUTO" if weight_type == "GS" else weight_type
        self.type_label.config(text=f"Tipo: {type_text}")

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        try:
            if hasattr(self, 'log_text') and self.log_text:
                self.log_text.insert(tk.END, log_entry)
                self.log_text.see(tk.END)
            else:
                print(log_entry, end='')
        except Exception as e:
            print(f"[{timestamp}] ERROR al escribir en log: {str(e)}")
            print(log_entry, end='')

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def save_log(self):
        try:
            filename = f"weight_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, 'w') as f:
                f.write(self.log_text.get(1.0, tk.END))
            self.log_message(f"Log guardado en {filename}")
        except Exception as e:
            self.log_message(f"Error al guardar: {str(e)}")

    def on_closing(self):
        if self.driver:
            self.driver.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = WeightMonitor(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()