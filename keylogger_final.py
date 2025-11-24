import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import warnings
import pandas as pd
import os
import threading
import wave
import pyaudio
import time
from pynput import keyboard, mouse

warnings.simplefilter(action='ignore', category=FutureWarning)

# Ruta donde se guardarán los datos
RUTA = "Monitorizaciones/"
df_teclas = pd.DataFrame(columns=["tecla", "datetime", "timestamp"])
df_mouse = pd.DataFrame(columns=["evento", "x", "y", "datetime", "timestamp"])
running = False  # Controla si el listener está activo
keyboard_thread = None  # Hilo para capturar el teclado
mouse_thread = None  # Hilo para capturar el ratón
audio_thread = None  # Hilo para capturar el audio
fecha_seleccionada = None  # ID del usuario
numero = 1  # Número de monitorización
recording = False  # Controla la grabación de audio
debug = True  # Modo de depuración
tiempo_inicio = None  # Tiempo de inicio de la grabación
tiempo_fin = None  # Tiempo de fin de la grabación

# Buffer y control de guardado para movimientos del ratón
mouse_move_buffer = []       # lista de dicts como las filas del DataFrame
last_mouse_move_flush = None # timestamp (float) del último guardado
MOUSE_MOVE_FLUSH_INTERVAL = 60.0  # segundos (1 minuto)

# Configuración del micrófono Poly 40
FORMAT = pyaudio.paInt16  # Formato de audio
CHANNELS = 1  # Mono
RATE = 44100  # Frecuencia de muestreo
CHUNK = 1024  # Tamaño del buffer
MICROPHONE_INDEX = None

CSV_FILE = "presentaciones_y_evaluaciones.csv"

if not os.path.exists(RUTA):
            os.makedirs(RUTA)
            
def _mouse_csv_path(fecha_seleccionada, numero):
    return os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_mouse.csv")

class KeyLoggerApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Captura de Teclado, Ratón y Audio")
        self.geometry("400x300")

        ttk.Label(self, text="Seleccione una fecha:").pack(pady=5)
        self.combo_fecha = ttk.Combobox(self, state="readonly", width=30)
        self.combo_fecha.pack(pady=5)
        self.load_csv()
        # Selección de micrófono
        self.label_mic = ttk.Label(self, text="Seleccione un micrófono:")
        self.label_mic.pack(pady=5)

        self.combo_mic = ttk.Combobox(self, state="readonly", width=40)
        self.combo_mic.pack(pady=5)
        self.load_microphones()

        # Botón para iniciar la captura
        self.start_button = ttk.Button(self, text="Iniciar Captura", command=self.start_logging)
        self.start_button.pack(pady=5)

        # Mensaje de estado
        self.status_label = ttk.Label(self, text="")
        self.status_label.pack(pady=10)

        # Botón para detener la captura
        self.stop_button = ttk.Button(self, text="Finalizar y Guardar", command=self.stop_logging, state=tk.DISABLED)
        self.stop_button.pack(pady=5)
    
    def _keyboard_csv_path(fecha_seleccionada, numero):
        return os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_keyboard.csv")

    def _mouse_csv_path(fecha_seleccionada, numero):
        return os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_mouse.csv")
    
    def load_csv(self):
        """Carga y ordena los datos del archivo CSV."""
        try:
            self.df = pd.read_csv(CSV_FILE)
            self.df["Fecha_presentacion"] = pd.to_datetime(self.df["Fecha_presentacion"], dayfirst=True)
            self.df = self.df.sort_values(by="Fecha_presentacion")
            self.df["Fecha_presentacion"] = self.df["Fecha_presentacion"].dt.strftime("%d/%m/%Y")
            self.combo_fecha["values"] = self.df["Fecha_presentacion"].unique().tolist()
        except FileNotFoundError:
            messagebox.showerror("Error", f"No se encontró el archivo {CSV_FILE}.")
            self.destroy()


    def load_microphones(self):
        """Carga los micrófonos disponibles en el combobox."""
        p = pyaudio.PyAudio()

        mic_devices = {}

        # Obtener la cantidad de dispositivos de audio
        info = p.get_host_api_info_by_index(0)
        numdevices = info.get('deviceCount')

        # Iterar sobre los dispositivos para encontrar los de entrada
        for i in range(0, numdevices):
            device_info = p.get_device_info_by_host_api_device_index(0, i)
            if device_info.get('maxInputChannels') > 0:  # Solo dispositivos de entrada
                mic_devices[device_info.get('name')] = i  # Guardar en el diccionario

        # Cerrar PyAudio
        p.terminate()

        self.mic_devices = mic_devices
        self.combo_mic["values"] = list(mic_devices.keys())
        if mic_devices:
            self.combo_mic.current(0)

    def start_logging(self):
        """Inicia la captura de teclado, ratón y audio en segundo plano."""
        global running, keyboard_thread, mouse_thread, audio_thread, fecha_seleccionada, numero, recording, tiempo_inicio, RAte, MICROPHONE_INDEX

        fecha_seleccionada = self.combo_fecha.get()
        if not fecha_seleccionada:
            messagebox.showerror("Error", "Debe seleccionar fecha.")
            return
        fecha_seleccionada = datetime.strptime(fecha_seleccionada, "%d/%m/%Y")
        fecha_seleccionada = fecha_seleccionada.strftime("%Y%m%d")
        numero = 1
        while os.path.exists(f"Monitorizaciones/{fecha_seleccionada}_{numero}_keyboard.csv"):
            numero += 1

        if not self.combo_mic.get():
            messagebox.showerror("Error", "Debe seleccionar un micrófono.")
            return

        MICROPHONE_INDEX = self.mic_devices[self.combo_mic.get()]

        # Detectar la tasa de muestreo automáticamente
        p = pyaudio.PyAudio()
        try:
            device_info = p.get_device_info_by_index(MICROPHONE_INDEX)
            RATE = int(device_info.get("defaultSampleRate", 44100))  # Si no encuentra la tasa, usa 44100
            print(f"Micrófono seleccionado: {device_info['name']} - Frecuencia de muestreo: {RATE} Hz")
        except Exception as e:
            print(f"Error detectando tasa de muestreo: {e}")
            RATE = 44100  # Valor por defecto en caso de error
        finally:
            p.terminate()

        running = True
        recording = True
        self.status_label.config(text=f"Capturando datos")
        self.start_button.config(state=tk.DISABLED)
        self.combo_fecha.config(state=tk.DISABLED)
        self.combo_mic.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        # Iniciar los listeners en hilos separados
        tiempo_inicio = datetime.now()
        
        global mouse_move_buffer, last_mouse_move_flush
        mouse_move_buffer = []
        last_mouse_move_flush = time.time()

        keyboard_thread = threading.Thread(target=self.start_key_listener, daemon=True)
        mouse_thread = threading.Thread(target=self.start_mouse_listener, daemon=True)
        #audio_thread = threading.Thread(target=self.start_audio_capture, daemon=True)

        keyboard_thread.start()
        mouse_thread.start()
        #audio_thread.start()

    def start_key_listener(self):
        """Inicia el listener de teclado en segundo plano."""
        with keyboard.Listener(on_press=self.capture_key) as listener:
            listener.join()

    def start_mouse_listener(self):
        """Inicia el listener de ratón en segundo plano."""
        with mouse.Listener(on_click=self.capture_mouse_click, on_move=self.capture_mouse_move) as listener:
            listener.join()
    
    def capture_key(self, key):
        """Captura la tecla presionada y la almacena en el DataFrame y en CSV."""
        global df_teclas, running, debug, fecha_seleccionada, numero
        if not running:
            return False  # Detiene el listener
        try:
            key_name = key.char  # Tecla normal (letras, números, etc.)
        except AttributeError:
            key_name = str(key)  # Teclas especiales (Shift, Ctrl, F1, etc.)
        timestamp = datetime.now()
        if debug:
            print(f"Tecla presionada: {key_name}")
        df_teclas.loc[len(df_teclas)] = {
            "tecla": key_name,
            "datetime": timestamp,
            "timestamp": timestamp.timestamp()
        }
        # Guardar inmediatamente en CSV
        if fecha_seleccionada:
            df_teclas.tail(1).to_csv(
                os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_keyboard.csv"),
                mode='a', header=not os.path.exists(os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_keyboard.csv")),
                index=False
            )

    def capture_mouse_click(self, x, y, button, pressed):
        """Registra los clics del ratón."""
        global df_mouse, running, debug

        if not running:
            return False

        evento = f"{'Presionado' if pressed else 'Liberado'} - {button}"
        timestamp = datetime.now()
        if debug:
            print(f"Ratón {evento} en ({x}, {y})")

        df_mouse.loc[len(df_mouse)] = {
                "evento": evento,
                "x": x,
                "y": y,
                "datetime": timestamp,
                "timestamp": timestamp.timestamp()
            }
        
        # Guardar inmediatamente en CSV (append de una sola fila)
        if fecha_seleccionada:
            df_mouse.tail(1).to_csv(
                os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_mouse.csv"),
                mode='a', header=not os.path.exists(os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_mouse.csv")),
                index=False
            )

    
    def capture_mouse_move(self, x, y):
        """Registra el movimiento del ratón en un buffer y guarda en CSV cada 1 minuto."""
        global df_mouse, running, fecha_seleccionada, numero
        global mouse_move_buffer, last_mouse_move_flush

        if not running:
            return False

        timestamp_dt = datetime.now()
        timestamp_unix = timestamp_dt.timestamp()

        row = {
            "evento": "Movimiento",
            "x": x,
            "y": y,
            "datetime": timestamp_dt,
            "timestamp": timestamp_unix
        }

        # Añadir al DataFrame (si quieres mantener df_mouse completo) y al buffer
        df_mouse.loc[len(df_mouse)] = row
        mouse_move_buffer.append(row)

        # Revisar si ha pasado el intervalo para hacer flush a CSV
        now = time.time()
        if last_mouse_move_flush is None:
            last_mouse_move_flush = now

        if (now - last_mouse_move_flush) >= MOUSE_MOVE_FLUSH_INTERVAL and mouse_move_buffer:
            # Volcar el buffer a CSV en bloque
            if fecha_seleccionada:
                csv_path = _mouse_csv_path(fecha_seleccionada, numero)
                pd.DataFrame(mouse_move_buffer).to_csv(
                    csv_path,
                    mode='a',
                    header=not os.path.exists(csv_path),
                    index=False
                )
            # Vaciar el buffer y actualizar el timestamp de último flush
            mouse_move_buffer.clear()
            last_mouse_move_flush = now


    def start_audio_capture(self):
        """Captura audio en fragmentos de 5 minutos desde el micrófono Poly 40."""
        global recording, fecha_seleccionada, numero, RATE, MICROPHONE_INDEX
        try:
            p = pyaudio.PyAudio()
            stream = p.open(format=FORMAT, channels=CHANNELS, rate=RATE, input=True, frames_per_buffer=CHUNK, input_device_index=MICROPHONE_INDEX)

            fragment_duration = 60 * 1  # 5 minutos en segundos
            start_time = time.time()
            frames = []
            fragment_counter = 1  # Contador para diferenciar los fragmentos
            fragment_files = []
            while recording:
                current_time = time.time()
                if current_time - start_time >= fragment_duration:
                    # Guardar el fragmento actual
                    audio_filename = os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_{fragment_counter}_audio.wav")
                    with wave.open(audio_filename, 'wb') as wf:
                        wf.setnchannels(CHANNELS)
                        wf.setsampwidth(p.get_sample_size(FORMAT))
                        wf.setframerate(RATE)
                        wf.writeframes(b''.join(frames))

                    print(f"Fragmento de audio guardado en {audio_filename}")
                    fragment_files.append(audio_filename)
                    frames = []  # Limpiar los frames para el siguiente fragmento
                    start_time = current_time  # Reiniciar el temporizador
                    fragment_counter += 1  # Incrementar el contador

                data = stream.read(CHUNK)
                frames.append(data)

            # Guardar el último fragmento cuando se detiene la grabación
            audio_filename = os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_{fragment_counter}_audio.wav")
            with wave.open(audio_filename, 'wb') as wf:
                wf.setnchannels(CHANNELS)
                wf.setsampwidth(p.get_sample_size(FORMAT))
                wf.setframerate(RATE)
                wf.writeframes(b''.join(frames))

            print(f"Audio guardado en {audio_filename}")
            fragment_files.append(audio_filename)

            stream.stop_stream()
            stream.close()
            p.terminate()

            audio_final = os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_audio_completo.wav")

            with wave.open(fragment_files[0], 'rb') as wf:
                params = wf.getparams()  # Obtener los parámetros del audio
                with wave.open(audio_final, 'wb') as audio_completo:
                    audio_completo.setparams(params)  # Usar los mismos parámetros

                    # Escribir todos los fragmentos al archivo final
                    for fragmento in fragment_files:
                        with wave.open(fragmento, 'rb') as wf:
                            audio_completo.writeframes(wf.readframes(wf.getnframes()))

            print(f"Audio completo guardado en {audio_final}")
        except Exception as e:
            print(f"Error en la captura de audio: {e}")
    
    def stop_logging(self):
        """Guarda los datos en archivos CSV y detiene la captura."""
        global df_teclas, df_mouse, running, recording, tiempo_inicio, tiempo_fin, numero, fecha_seleccionada, keyboard_thread, mouse_thread, audio_thread

        running = False
        recording = False  # Detiene la grabación de audio
        tiempo_fin = datetime.now()

        # Flush final del buffer de movimientos (si quedó algo pendiente)
        global mouse_move_buffer, last_mouse_move_flush, fecha_seleccionada, numero
        if mouse_move_buffer and fecha_seleccionada:
            csv_path = _mouse_csv_path(fecha_seleccionada, numero)
            pd.DataFrame(mouse_move_buffer).to_csv(
                csv_path,
                mode='a',
                header=not os.path.exists(csv_path),
                index=False
            )
            mouse_move_buffer.clear()

        df = pd.DataFrame(columns=["Inicio","Inicio_UNIX","Fin","Fin_UNIX"])
        df.loc[len(df)] = {
                "Inicio": tiempo_inicio,
                "Inicio_UNIX": tiempo_inicio.timestamp(),
                "Fin": tiempo_fin,
                "Fin_UNIX": tiempo_fin.timestamp()
            }
            
        df.to_csv(os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_time.csv"), index=False)

        df_teclas.to_csv(os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_keyboard.csv"), index=False)

        df_mouse.to_csv(os.path.join(RUTA, f"{fecha_seleccionada}_{numero}_mouse.csv"), index=False)

        messagebox.showinfo("Guardado", "Datos guardados correctamente.")
        self.destroy()

if __name__ == "__main__":
    app = KeyLoggerApp()
    app.mainloop()