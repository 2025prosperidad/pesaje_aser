import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import serial
import serial.tools.list_ports
import threading
import re
from datetime import datetime
import time
import os
import psutil

# Intentar importar tkwebview2 (solo disponible en Windows)
USE_WEBVIEW2 = False
try:
    from tkwebview2.tkwebview2 import WebView2, have_runtime, install_runtime
    USE_WEBVIEW2 = True
except ImportError:
    pass

# Fallback: Selenium para sistemas sin WebView2
USE_SELENIUM = False
if not USE_WEBVIEW2:
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        from PIL import Image, ImageTk
        import io
        USE_SELENIUM = True
    except ImportError:
        pass

HIK_CONNECT_URL = "https://www.hik-connect.com/views/login/index.html#/portal"

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
        self.webview = None
        self.driver = None

        self.setup_ui()
        self.log_message("=== INICIANDO APLICACIÓN ===")

        if USE_WEBVIEW2:
            self.log_message("✓ Usando WebView2 (tiempo real)")
            self.root.after(500, self.init_webview2)
        elif USE_SELENIUM:
            self.log_message("⚠ WebView2 no disponible, usando Selenium (fallback)")
            self.root.after(1000, self.init_selenium)
        else:
            self.log_message("❌ No hay motor de navegador disponible")

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

        self.btn_refresh = tk.Button(config_frame, text="🔄", command=self.refresh_ports, bg="#3d3d3d", fg="white", width=3)
        self.btn_refresh.grid(row=0, column=4, padx=2)

        # NUEVO: Botón para liberar puerto específico
        self.btn_force_free = tk.Button(config_frame, text="🔓 Liberar", command=self.force_free_selected_port,
                                       bg="#9c27b0", fg="white", width=10, font=("Arial", 9, "bold"))
        self.btn_force_free.grid(row=0, column=5, padx=5)

        # Botón para reiniciar puertos
        self.btn_reset_ports = tk.Button(config_frame, text="⚡ Reiniciar", command=self.reset_ports,
                                        bg="#ff6600", fg="white", width=10, font=("Arial", 9, "bold"))
        self.btn_reset_ports.grid(row=0, column=6, padx=5)

        self.btn_connect = tk.Button(config_frame, text="Conectar", command=self.toggle_connection,
                                     bg="#0d7377", fg="white", width=12, font=("Arial", 10, "bold"))
        self.btn_connect.grid(row=0, column=7, padx=10)

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
        self.right_container = tk.Frame(main_horizontal_frame, bg="#2d2d2d")
        self.right_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Título
        title_frame = tk.Frame(self.right_container, bg="#2d2d2d", height=30)
        title_frame.pack(fill=tk.X, side=tk.TOP)
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text="Hik-Connect - Tiempo Real",
                bg="#2d2d2d", fg="white", font=("Arial", 11, "bold")).pack(pady=5)

        # Contenedor del navegador (se llenará con WebView2 o Canvas según disponibilidad)
        self.browser_container = tk.Frame(self.right_container, bg="black")
        self.browser_container.pack(fill=tk.BOTH, expand=True)

        # Botones inferiores
        btn_frame = tk.Frame(self.root, bg="#1e1e1e")
        btn_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(btn_frame, text="Limpiar Log", command=self.clear_log,
                 bg="#3d3d3d", fg="white").pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Guardar Datos", command=self.save_log,
                 bg="#3d3d3d", fg="white").pack(side=tk.LEFT, padx=5)

    # =============================================
    # WebView2 - Navegador embebido en tiempo real
    # =============================================
    def init_webview2(self):
        """Inicializa el navegador WebView2 embebido (solo Windows)"""
        try:
            self.log_message("🌐 Iniciando navegador WebView2...")

            # Verificar si el runtime de WebView2 está instalado
            if not have_runtime():
                self.log_message("⏬ Instalando WebView2 Runtime...")
                install_runtime()
                self.log_message("✓ WebView2 Runtime instalado")

            # Forzar actualización para obtener dimensiones
            self.root.update_idletasks()
            w = self.browser_container.winfo_width()
            h = self.browser_container.winfo_height()
            if w <= 1:
                w = 800
            if h <= 1:
                h = 600

            # Crear widget WebView2 dentro del contenedor
            self.webview = WebView2(self.browser_container, width=w, height=h, url=HIK_CONNECT_URL)
            self.webview.pack(fill=tk.BOTH, expand=True)

            self.log_message("✓ Navegador WebView2 iniciado correctamente (tiempo real)")

        except Exception as e:
            self.log_message(f"❌ Error al iniciar WebView2: {e}")
            # Mostrar error en la interfaz
            error_label = tk.Label(self.browser_container,
                text=f"Error al iniciar navegador WebView2:\n{str(e)}\n\nInstale tkwebview2: pip install tkwebview2",
                bg="black", fg="red", font=("Arial", 12), wraplength=500)
            error_label.pack(expand=True)

    # =============================================
    # Selenium - Fallback para sistemas sin WebView2
    # =============================================
    def init_selenium(self):
        try:
            self.log_message("🌐 Iniciando navegador Chrome (Selenium fallback)...")
            chrome_options = Options()

            # Crear carpeta temporal para perfil de Chrome
            temp_profile = os.path.join(os.getcwd(), "chrome_profile")
            if not os.path.exists(temp_profile):
                os.makedirs(temp_profile)

            # Configurar Chrome con perfil temporal persistente
            chrome_options.add_argument(f"--user-data-dir={temp_profile}")
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)

            self.log_message(f"✓ Usando perfil: {temp_profile}")

            # Crear Canvas para mostrar screenshots (fallback)
            self.browser_canvas = tk.Canvas(self.browser_container, bg="black", highlightthickness=0)
            self.browser_canvas.pack(fill=tk.BOTH, expand=True)

            # Usar webdriver-manager para descargar automáticamente ChromeDriver
            self.driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=chrome_options
            )
            self.driver.get(HIK_CONNECT_URL)

            self.log_message("✓ Navegador Chrome iniciado (modo screenshot cada 1s)")
            self.update_browser_view()

        except Exception as e:
            self.log_message(f"❌ Error al iniciar Selenium: {e}")
            if not hasattr(self, 'browser_canvas'):
                self.browser_canvas = tk.Canvas(self.browser_container, bg="black", highlightthickness=0)
                self.browser_canvas.pack(fill=tk.BOTH, expand=True)
            self.browser_canvas.delete("all")
            self.browser_canvas.create_text(
                self.browser_canvas.winfo_width()//2,
                self.browser_canvas.winfo_height()//2,
                text=f"Error al iniciar navegador:\n{str(e)}",
                fill="red",
                font=("Arial", 12)
            )

    def update_browser_view(self):
        """Actualización por screenshots (solo se usa con Selenium fallback)"""
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
            self.log_message(f"⚠ Error al actualizar vista del navegador: {e}")

    # =============================================
    # Puertos seriales
    # =============================================
    def force_free_selected_port(self):
        """Liberar específicamente el puerto seleccionado de forma AGRESIVA"""
        port = self.port_combo.get()
        if not port:
            messagebox.showwarning("Sin Puerto", "Selecciona primero un puerto COM")
            return

        try:
            self.log_message(f"🔓 ═══ LIBERACIÓN AGRESIVA DE {port} ═══")

            # PASO 1: Desconectar si está conectado
            if self.is_running:
                self.log_message("⏹ Deteniendo conexión activa...")
                self.disconnect()
                time.sleep(0.5)

            # PASO 2: Buscar y MATAR todos los procesos Python
            self.log_message("🔨 Buscando procesos que bloquean el puerto...")
            current_pid = os.getpid()
            killed_processes = []

            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    if proc.info['pid'] == current_pid:
                        continue

                    proc_name = proc.info['name'].lower()

                    # Matar TODOS los procesos Python excepto este
                    if 'python' in proc_name:
                        try:
                            proc.kill()
                            killed_processes.append(f"{proc.info['name']} (PID: {proc.info['pid']})")
                            self.log_message(f"💀 Proceso eliminado: {proc.info['name']} (PID: {proc.info['pid']})")
                            time.sleep(0.2)
                        except (psutil.NoSuchProcess, psutil.AccessDenied):
                            pass

                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass

            time.sleep(1)

            # PASO 3: Intentar forzar apertura/cierre del puerto
            self.log_message(f"🔄 Forzando ciclos de apertura/cierre en {port}...")
            freed = False
            for i in range(10):  # 10 intentos
                try:
                    s = serial.Serial(port, baudrate=9600, timeout=0.1)
                    s.close()
                    del s
                    freed = True
                    self.log_message(f"✓ Ciclo {i+1}/10: Puerto respondió")
                    time.sleep(0.1)
                except Exception as e:
                    if i == 9:
                        self.log_message(f"❌ Ciclo {i+1}/10: Falló - {str(e)[:30]}")
                    time.sleep(0.1)

            # PASO 4: Actualizar lista de puertos
            self.refresh_ports()

            # Resultado
            if killed_processes or freed:
                result_msg = f"✅ Puerto {port} liberado!\n\n"
                if killed_processes:
                    result_msg += f"Procesos eliminados: {len(killed_processes)}\n"
                    for p in killed_processes[:3]:  # Mostrar máximo 3
                        result_msg += f"• {p}\n"
                if freed:
                    result_msg += f"\n✓ Puerto respondió correctamente"

                self.log_message(f"✅ Liberación completada: {len(killed_processes)} procesos eliminados")
                messagebox.showinfo("Puerto Liberado", result_msg + "\n\nAhora intenta CONECTAR")
            else:
                self.log_message("⚠ No se pudo liberar el puerto automáticamente")
                messagebox.showwarning("Liberación Manual",
                    f"No se liberó el puerto automáticamente.\n\n"
                    f"ACCIÓN REQUERIDA:\n\n"
                    f"1. DESCONECTA el cable USB\n"
                    f"2. Espera 5 segundos\n"
                    f"3. RECONECTA el cable USB\n"
                    f"4. Presiona 🔄 para actualizar\n"
                    f"5. Intenta Conectar nuevamente")

        except Exception as e:
            self.log_message(f"❌ Error crítico al liberar puerto: {str(e)}")
            messagebox.showerror("Error Crítico",
                f"Error al liberar puerto:\n{str(e)}\n\n"
                f"SOLUCIÓN:\n"
                f"1. Desconecta el cable USB\n"
                f"2. Cierra ESTE programa (X)\n"
                f"3. Reconecta el cable USB\n"
                f"4. Abre el programa de nuevo")

    def refresh_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo['values'] = ports
        if ports:
            if 'COM5' in ports:
                self.port_combo.set('COM5')
            else:
                self.port_combo.set(ports[0])
        self.log_message(f"Puertos disponibles: {', '.join(ports) if ports else 'ninguno'}")

    def reset_ports(self):
        """Reinicia todos los puertos COM cerrando procesos que los usan"""
        try:
            # Primero desconectar si está conectado
            if self.is_running:
                self.disconnect()
                time.sleep(0.5)

            self.log_message("🔄 Reiniciando puertos COM...")

            # Forzar cierre de cualquier conexión serial abierta en todos los puertos COM
            import serial.tools.list_ports
            ports = [port.device for port in serial.tools.list_ports.comports()]

            freed_count = 0
            for port in ports:
                try:
                    # Intentar abrir y cerrar cada puerto para liberarlo
                    temp_serial = serial.Serial(port, baudrate=9600, timeout=0.1)
                    temp_serial.close()
                    freed_count += 1
                    self.log_message(f"✓ Puerto {port} liberado")
                    time.sleep(0.1)
                except Exception as e:
                    self.log_message(f"⚠ {port}: {str(e)[:50]}")

            time.sleep(0.5)

            # Cerrar todos los procesos que puedan estar usando puertos COM
            current_pid = os.getpid()
            killed_count = 0

            try:
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        # No matar el proceso actual
                        if proc.info['pid'] == current_pid:
                            continue

                        proc_name = proc.info['name'].lower()

                        # Buscar procesos relacionados con puertos seriales
                        if any(keyword in proc_name for keyword in ['python', 'serial', 'arduino', 'putty', 'terminal', 'com']):
                            try:
                                # Verificar archivos abiertos por el proceso
                                has_com_port = False
                                try:
                                    for item in proc.open_files():
                                        if 'COM' in item.path.upper() or '\\Device\\Serial' in item.path:
                                            has_com_port = True
                                            break
                                except (psutil.AccessDenied, psutil.NoSuchProcess):
                                    pass

                                # Si tiene puerto COM abierto, terminarlo
                                if has_com_port:
                                    proc.terminate()
                                    try:
                                        proc.wait(timeout=2)
                                    except psutil.TimeoutExpired:
                                        proc.kill()
                                    killed_count += 1
                                    self.log_message(f"✓ Proceso terminado: {proc.info['name']} (PID: {proc.info['pid']})")

                            except (psutil.AccessDenied, psutil.NoSuchProcess, psutil.TimeoutExpired):
                                pass

                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        pass
            except Exception as e:
                self.log_message(f"⚠ Error al buscar procesos: {str(e)[:50]}")

            time.sleep(1)

            # Refrescar lista de puertos
            self.refresh_ports()

            total_actions = freed_count + killed_count

            if total_actions > 0:
                self.log_message(f"✓ Puertos reiniciados: {freed_count} liberado(s), {killed_count} proceso(s) terminado(s).")
                messagebox.showinfo("Puertos Reiniciados",
                                   f"Resultado:\n"
                                   f"• {freed_count} puerto(s) liberado(s)\n"
                                   f"• {killed_count} proceso(s) terminado(s)\n\n"
                                   f"Intenta conectar nuevamente.")
            else:
                self.log_message("ℹ No se pudo liberar ningún puerto automáticamente.")
                response = messagebox.askyesno("Acción Manual Requerida",
                                   "No se pudieron liberar los puertos automáticamente.\n\n"
                                   "SOLUCIÓN:\n"
                                   "1. Desconecta el cable USB del dispositivo\n"
                                   "2. Espera 5 segundos\n"
                                   "3. Reconecta el cable USB\n\n"
                                   "¿Deseas abrir el Administrador de Tareas para\n"
                                   "cerrar procesos manualmente?")
                if response:
                    os.system('taskmgr')

        except Exception as e:
            self.log_message(f"❌ Error al reiniciar puertos: {str(e)}")
            messagebox.showerror("Error", f"Error al reiniciar puertos:\n{str(e)}")

    def toggle_connection(self):
        if not self.is_running:
            self.connect()
        else:
            self.disconnect()

    def force_close_port(self, port):
        """Forzar cierre de un puerto específico"""
        try:
            self.log_message(f"🔨 Forzando cierre de {port}...")

            # Método 1: Cerrar si hay puerto serial activo en esta instancia
            if self.serial_port:
                try:
                    if self.serial_port.is_open:
                        self.serial_port.close()
                    self.serial_port = None
                    self.log_message(f"✓ Puerto de instancia cerrado")
                    time.sleep(0.3)
                except:
                    pass

            # Método 2: Intentar abrir y cerrar múltiples veces
            for i in range(3):
                try:
                    s = serial.Serial(port, timeout=0.1)
                    s.close()
                    del s
                    time.sleep(0.2)
                    self.log_message(f"✓ Ciclo de apertura/cierre {i+1} exitoso")
                except:
                    pass

            # Método 3: Buscar procesos usando el puerto (excepto este)
            current_pid = os.getpid()
            killed = False

            try:
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        if proc.info['pid'] == current_pid:
                            continue

                        proc_name = proc.info['name'].lower()
                        # Solo buscar procesos relacionados con serial
                        if 'python' not in proc_name:
                            continue

                        # Verificar archivos abiertos
                        try:
                            for item in proc.open_files():
                                if port.upper() in item.path.upper():
                                    self.log_message(f"🔨 Matando proceso {proc.info['name']} (PID: {proc.info['pid']})")
                                    proc.kill()
                                    time.sleep(0.5)
                                    killed = True
                                    break
                        except (psutil.AccessDenied, AttributeError):
                            pass

                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
            except Exception as e:
                self.log_message(f"⚠ Error buscando procesos: {str(e)[:40]}")

            if killed:
                time.sleep(0.5)
                self.log_message(f"✓ Puerto {port} debería estar libre ahora")

            return True

        except Exception as e:
            self.log_message(f"⚠ Error en force_close_port: {str(e)[:50]}")
            return False

    def connect(self):
        try:
            port = self.port_combo.get()

            if not port:
                messagebox.showwarning("Sin Puerto", "Por favor seleccione un puerto COM")
                return

            baud = int(self.baud_combo.get())

            # PASO 1: Asegurarse de que no hay conexión previa
            self.log_message(f"🔌 Preparando conexión a {port}...")
            if self.serial_port:
                try:
                    if self.serial_port.is_open:
                        self.serial_port.close()
                    self.serial_port = None
                    time.sleep(0.3)
                except:
                    pass

            # PASO 2: UN SOLO intento de conexión directa
            self.log_message(f"📡 Intentando abrir {port} @ {baud} baud...")

            self.serial_port = serial.Serial(
                port=port,
                baudrate=baud,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1,
                write_timeout=1
            )

            # Verificar que se abrió correctamente
            if not self.serial_port.is_open:
                raise serial.SerialException("El puerto no se abrió correctamente")

            self.is_running = True
            self.btn_connect.config(text="Desconectar", bg="#d32f2f")
            self.status_label.config(fg="#00ff00")
            self.status_text.config(text="CONECTADO", fg="#00ff00")

            # Iniciar thread de lectura
            self.read_thread = threading.Thread(target=self.read_serial, daemon=True)
            self.read_thread.start()

            self.log_message(f"✅ ¡CONECTADO EXITOSAMENTE a {port} @ {baud} baud!")

        except serial.SerialException as e:
            error_msg = str(e)
            self.log_message(f"❌ Error de conexión: {error_msg}")

            # Limpiar cualquier referencia
            if self.serial_port:
                try:
                    self.serial_port.close()
                except:
                    pass
                self.serial_port = None

            # NO PREGUNTAR SI REINTENTAR - Mostrar opciones claras
            messagebox.showerror("Puerto Bloqueado",
                f"❌ NO SE PUDO CONECTAR A {port}\n\n"
                f"El puerto está siendo usado por otro programa.\n\n"
                f"SOLUCIONES (en orden):\n\n"
                f"1️⃣ DESCONECTA el cable USB por 5 segundos\n"
                f"    y vuélvelo a conectar\n\n"
                f"2️⃣ Haz clic en '🔓 Liberar' y luego 'Conectar'\n\n"
                f"3️⃣ Cierra este programa completamente (X)\n"
                f"    y ábrelo de nuevo\n\n"
                f"4️⃣ Reinicia Windows si nada funciona")

        except Exception as e:
            self.log_message(f"❌ Error inesperado: {str(e)}")
            if self.serial_port:
                try:
                    self.serial_port.close()
                except:
                    pass
                self.serial_port = None
            messagebox.showerror("Error", f"Error inesperado:\n{str(e)}")

    def disconnect(self):
        self.is_running = False

        # Esperar a que el thread de lectura termine
        if hasattr(self, 'read_thread') and self.read_thread and self.read_thread.is_alive():
            self.log_message("⏳ Esperando cierre del thread de lectura...")
            self.read_thread.join(timeout=2)

        # Cerrar puerto con múltiples intentos
        if self.serial_port:
            for attempt in range(3):
                try:
                    if self.serial_port.is_open:
                        self.serial_port.close()
                    self.log_message(f"✓ Puerto cerrado correctamente (intento {attempt + 1})")
                    break
                except Exception as e:
                    if attempt < 2:
                        self.log_message(f"⚠ Intento {attempt + 1} de cierre falló, reintentando...")
                        time.sleep(0.2)
                    else:
                        self.log_message(f"❌ Error al cerrar puerto: {str(e)}")

            # Liberar la referencia
            self.serial_port = None
            time.sleep(0.3)

        self.btn_connect.config(text="Conectar", bg="#0d7377")
        self.status_label.config(fg="#ff4444")
        self.status_text.config(text="DESCONECTADO", fg="#888888")
        self.log_message("═══ DESCONECTADO ═══")

    def read_serial(self):
        while self.is_running:
            try:
                if self.serial_port and self.serial_port.is_open:
                    line = self.serial_port.readline().decode('utf-8', errors='ignore').strip()
                    if line:
                        self.root.after(0, self.process_data, line)
            except Exception as e:
                self.log_message(f"❌ Error de lectura: {str(e)}")
                break

    def process_data(self, data):
        self.log_message(f"📊 Datos: {data}")
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
        self.log_message("🗑 Log limpiado")

    def save_log(self):
        try:
            filename = f"weight_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get(1.0, tk.END))
            self.log_message(f"💾 Log guardado en {filename}")
            messagebox.showinfo("Guardado", f"Log guardado exitosamente en:\n{filename}")
        except Exception as e:
            self.log_message(f"❌ Error al guardar: {str(e)}")
            messagebox.showerror("Error", f"Error al guardar log:\n{str(e)}")

    def cleanup_all_connections(self):
        """Limpia todas las conexiones al cerrar la aplicación"""
        self.log_message("🔄 Cerrando aplicación y liberando recursos...")

        # Desconectar puerto serial
        if self.is_running:
            self.disconnect()

        # Cerrar navegador Selenium si se usó como fallback
        if self.driver:
            try:
                self.driver.quit()
                self.log_message("✓ Navegador Selenium cerrado")
            except:
                pass

        # Pequeña pausa para asegurar que todo se cierre
        time.sleep(0.5)

    def on_closing(self):
        """Maneja el cierre de la aplicación"""
        if messagebox.askokcancel("Salir", "¿Desea cerrar la aplicación?"):
            self.cleanup_all_connections()
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = WeightMonitor(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
