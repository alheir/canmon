import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import serial
import serial.tools.list_ports
import threading
import time
import re
import platform
import os
from datetime import datetime

class CanMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TP2 CAN Monitor")
        self.root.geometry("1000x700")
        
        # Variables
        self.serial_port = None
        self.is_connected = False
        self.reading_thread = None
        self.should_read = False
        self.port_info = {}  # Almacenará información detallada de los puertos
        self.last_update_times = {}  # Almacenará timestamps de actualizaciones
        self.update_timer = None  # Para la actualización periódica de timestamps
        
        # Crear interfaz
        self.create_widgets()
        
        # Actualizar lista de puertos COM
        self.refresh_ports()
    
    def create_widgets(self):
        # Marco principal con dos columnas
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === COLUMNA IZQUIERDA ===
        left_frame = ttk.LabelFrame(main_frame, text="Conexión y Control", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Sección de conexión serial
        conn_frame = ttk.Frame(left_frame)
        conn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(conn_frame, text="Puerto:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.port_combo = ttk.Combobox(conn_frame, width=25)
        self.port_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.port_combo.bind('<<ComboboxSelected>>', self.on_port_selected)
        
        self.refresh_btn = ttk.Button(conn_frame, text="Actualizar", command=self.refresh_ports)
        self.refresh_btn.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        
        self.connect_btn = ttk.Button(conn_frame, text="Conectar", command=self.toggle_connection)
        self.connect_btn.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        # Información de puerto
        self.port_info_label = ttk.Label(conn_frame, text="", wraplength=300)
        self.port_info_label.grid(row=1, column=0, columnspan=4, sticky=tk.W, padx=5, pady=5)
        
        # Sección de envío de mensajes CAN personalizados
        send_frame = ttk.LabelFrame(left_frame, text="Enviar mensaje CAN", padding=10)
        send_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(send_frame, text="ID (hex):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.can_id_entry = ttk.Entry(send_frame, width=5)
        self.can_id_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.can_id_entry.insert(0, "100")
        
        ttk.Label(send_frame, text="Datos (hex):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.can_data_entry = ttk.Entry(send_frame, width=20)
        self.can_data_entry.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        self.send_btn = ttk.Button(send_frame, text="Enviar", command=self.send_can_message)
        self.send_btn.grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)
        
        # Sección de presets para TP2
        presets_frame = ttk.LabelFrame(left_frame, text="Presets TP2", padding=10)
        presets_frame.pack(fill=tk.X, pady=10)
        
        # Selección de grupo
        ttk.Label(presets_frame, text="Grupo:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.group_combo = ttk.Combobox(presets_frame, width=5, values=[f"{i}" for i in range(8)])
        self.group_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.group_combo.current(0)
        self.group_combo.bind("<<ComboboxSelected>>", self.on_group_selected)
        
        # Tipo de ángulo
        ttk.Label(presets_frame, text="Tipo:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.angle_type = tk.StringVar(value="R")
        ttk.Radiobutton(presets_frame, text="Rolido (R)", variable=self.angle_type, value="R").grid(row=1, column=1)
        ttk.Radiobutton(presets_frame, text="Cabeceo (C)", variable=self.angle_type, value="C").grid(row=1, column=2)
        ttk.Radiobutton(presets_frame, text="Orientación (O)", variable=self.angle_type, value="O").grid(row=1, column=3)
        
        # Valor del ángulo
        ttk.Label(presets_frame, text="Ángulo:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.angle_value = ttk.Spinbox(presets_frame, from_=-179, to=180, width=5)
        self.angle_value.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        self.angle_value.set("0")
        
        self.send_preset_btn = ttk.Button(presets_frame, text="Enviar ángulo", command=self.send_tp2_angle)
        self.send_preset_btn.grid(row=2, column=3, sticky=tk.W, padx=5, pady=5)
        
        # Modo CAN
        mode_frame = ttk.LabelFrame(left_frame, text="Modo CAN", padding=10)
        mode_frame.pack(fill=tk.X, pady=10)
        
        self.normal_mode_btn = ttk.Button(mode_frame, text="Modo Normal", command=lambda: self.set_can_mode("NORMAL"))
        self.normal_mode_btn.grid(row=0, column=0, padx=5, pady=5)
        
        self.loopback_mode_btn = ttk.Button(mode_frame, text="Modo Loopback", command=lambda: self.set_can_mode("LOOPBACK"))
        self.loopback_mode_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # === COLUMNA DERECHA ===
        right_frame = ttk.LabelFrame(main_frame, text="Mensajes CAN", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Área para mostrar mensajes recibidos
        self.rx_text = scrolledtext.ScrolledText(right_frame, width=50, height=20)
        self.rx_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Configurar colores para los mensajes
        self.rx_text.tag_config("tx_msg", foreground="green")
        self.rx_text.tag_config("rx_msg", foreground="blue")
        self.rx_text.tag_config("system", foreground="black")
        self.rx_text.tag_config("error", foreground="red")
        
        # Botones para limpiar y habilitar/deshabilitar autoscroll
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.clear_btn = ttk.Button(btn_frame, text="Limpiar", command=lambda: self.rx_text.delete(1.0, tk.END))
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Área para mensajes TP2 interpretados
        tp2_frame = ttk.LabelFrame(right_frame, text="Mensajes TP2 Interpretados", padding=10)
        tp2_frame.pack(fill=tk.X, pady=10)
        
        # Tabla para mostrar ángulos
        columns = ('grupo', 'rolido', 'rolido_tiempo', 'cabeceo', 'cabeceo_tiempo', 'orientacion', 'orientacion_tiempo', 'ultima_act')
        self.tp2_tree = ttk.Treeview(tp2_frame, columns=columns, show='headings', height=8)
        
        # Definir encabezados
        self.tp2_tree.heading('grupo', text='Grupo')
        self.tp2_tree.heading('rolido', text='Rolido')
        self.tp2_tree.heading('rolido_tiempo', text='Tiempo R')
        self.tp2_tree.heading('cabeceo', text='Cabeceo')
        self.tp2_tree.heading('cabeceo_tiempo', text='Tiempo C')
        self.tp2_tree.heading('orientacion', text='Orientación')
        self.tp2_tree.heading('orientacion_tiempo', text='Tiempo O')
        self.tp2_tree.heading('ultima_act', text='Última Act.')
        
        # Ajustar anchos de columnas
        self.tp2_tree.column('grupo', width=50, anchor=tk.CENTER)
        self.tp2_tree.column('rolido', width=60, anchor=tk.CENTER)
        self.tp2_tree.column('rolido_tiempo', width=80, anchor=tk.CENTER)
        self.tp2_tree.column('cabeceo', width=60, anchor=tk.CENTER)
        self.tp2_tree.column('cabeceo_tiempo', width=80, anchor=tk.CENTER)
        self.tp2_tree.column('orientacion', width=60, anchor=tk.CENTER)
        self.tp2_tree.column('orientacion_tiempo', width=80, anchor=tk.CENTER)
        self.tp2_tree.column('ultima_act', width=100, anchor=tk.CENTER)
        
        # Estilos para la tabla
        self.tp2_tree.tag_configure('stale', foreground='gray')
        self.tp2_tree.tag_configure('active', foreground='black')
        
        self.tp2_tree.pack(fill=tk.BOTH, expand=True)
        
        # Inicializar con grupos del 0 al 7 (según TP2)
        for i in range(8):
            self.tp2_tree.insert('', tk.END, values=(i, '--', 'Nunca', '--', 'Nunca', '--', 'Nunca', 'Nunca'), tags=('stale',))
            self.last_update_times[i] = {
                'R': None,
                'C': None,
                'O': None,
                'any': None
            }

    def on_port_selected(self, event):
        """Muestra información detallada del puerto seleccionado"""
        selected = self.port_combo.get()
        if selected in self.port_info:
            info = self.port_info[selected]
            info_text = f"Puerto: {info['device']}\n"
            if info['description']:
                info_text += f"Descripción: {info['description']}\n"
            if info['manufacturer']:
                info_text += f"Fabricante: {info['manufacturer']}\n"
            if info['hwid']:
                info_text += f"Hardware ID: {info['hwid']}\n"
            if info['serial_number'] and info['serial_number'] != 'None':
                info_text += f"Número de serie: {info['serial_number']}\n"
            
            self.port_info_label.config(text=info_text)
        else:
            self.port_info_label.config(text="")
    
    def on_group_selected(self, event):
        """Actualiza la entrada de ID CAN cuando se selecciona un grupo"""
        group_id = int(self.group_combo.get())
        can_id = f"{0x100 + group_id:x}"
        self.can_id_entry.delete(0, tk.END)
        self.can_id_entry.insert(0, can_id)
    
    def refresh_ports(self):
        """Actualiza la lista de puertos serie disponibles con información detallada"""
        self.port_info = {}
        ports = []
        display_names = []
        
        try:
            for port in serial.tools.list_ports.comports():
                # Crear un identificador único para el puerto
                port_id = port.device
                
                # Guardar información detallada del puerto
                self.port_info[port_id] = {
                    'device': port.device,
                    'name': port.name if hasattr(port, 'name') else '',
                    'description': port.description if hasattr(port, 'description') else '',
                    'hwid': port.hwid if hasattr(port, 'hwid') else '',
                    'vid': port.vid if hasattr(port, 'vid') else None,
                    'pid': port.pid if hasattr(port, 'pid') else None,
                    'serial_number': port.serial_number if hasattr(port, 'serial_number') else '',
                    'manufacturer': port.manufacturer if hasattr(port, 'manufacturer') else '',
                    'product': port.product if hasattr(port, 'product') else '',
                    'interface': port.interface if hasattr(port, 'interface') else '',
                }
                
                # Crear un nombre descriptivo para mostrar en el ComboBox
                display_name = port.device
                if port.description and port.description != port.device:
                    display_name = f"{port.device} - {port.description}"
                
                display_names.append(display_name)
                ports.append(port_id)
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al detectar puertos: {str(e)}")
        
        # Actualizar el ComboBox
        self.port_combo['values'] = display_names
        if display_names:
            self.port_combo.current(0)
            self.on_port_selected(None)  # Mostrar info del primer puerto
        else:
            self.port_info_label.config(text="No se detectaron puertos serie")
    
    def toggle_connection(self):
        """Conecta o desconecta del puerto serial"""
        if not self.is_connected:
            selected = self.port_combo.get()
            
            # Extraer el nombre del dispositivo del texto mostrado (puede contener descripción)
            device = selected.split(' - ')[0] if ' - ' in selected else selected
            
            # Si está en el diccionario de info, usar ese puerto
            if device in self.port_info:
                port = self.port_info[device]['device']
            else:
                # Si no está en el diccionario, usar lo seleccionado directamente
                port = device
                
            try:
                self.serial_port = serial.Serial(port, 115200, timeout=1)
                self.is_connected = True
                self.connect_btn['text'] = "Desconectar"
                self.should_read = True
                
                # Restablecer datos TP2 al conectar
                self.reset_tp2_data()
                
                # Iniciar hilo para lectura continua
                self.reading_thread = threading.Thread(target=self.read_serial_data)
                self.reading_thread.daemon = True
                self.reading_thread.start()
                
                # Iniciar actualización periódica de timestamps
                self.start_timestamp_updates()
                
                # Mostrar información del sistema
                os_info = platform.platform()
                self.rx_text.insert(tk.END, f"Sistema: {os_info}\n", "system")
                self.rx_text.insert(tk.END, f"Conectado a {port} @ 115200 bps\n", "system")
                self.rx_text.see(tk.END)
            except Exception as e:
                messagebox.showerror("Error de conexión", str(e))
        else:
            self.should_read = False
            if self.serial_port:
                self.serial_port.close()
            self.is_connected = False
            self.connect_btn['text'] = "Conectar"
            self.rx_text.insert(tk.END, "Desconectado\n", "system")
            self.rx_text.see(tk.END)
            
            # Restablecer datos TP2 al desconectar
            self.reset_tp2_data()
            
            # Detener actualización de timestamps
            self.stop_timestamp_updates()
    
    def read_serial_data(self):
        """Lee datos del puerto serial continuamente"""
        while self.should_read:
            if self.serial_port and self.serial_port.in_waiting:
                try:
                    data = self.serial_port.readline().decode('utf-8').strip()
                    self.process_received_data(data)
                except Exception as e:
                    self.root.after(0, lambda: self.rx_text.insert(tk.END, f"Error de lectura: {str(e)}\n", "error"))
                    self.root.after(0, lambda: self.rx_text.see(tk.END))
            time.sleep(0.01)
    
    def process_received_data(self, data):
        """Procesa los datos recibidos por serial"""
        if not data:
            return
        
        # Añadir al área de texto con color azul para recepción
        self.root.after(0, lambda: self.rx_text.insert(tk.END, f"{data}\n", "rx_msg"))
        self.root.after(0, lambda: self.rx_text.see(tk.END))
        
        # Comprobar si es un mensaje CAN del formato TP2
        if data.startswith("CAN_RX_"):
            try:
                # Formato esperado: CAN_RX_ID_LEN_BYTE1_BYTE2_..._TP2_TYPE_VALUE
                parts = data.split("_")
                
                if len(parts) >= 5:
                    # Extraer ID para determinar el grupo
                    id_hex = parts[2]
                    try:
                        group_id = int(id_hex, 16) - 0x100
                    except ValueError:
                        print(f"Error al convertir ID: {id_hex}")
                        return
                        
                    # Verificar si es un ID dentro del rango TP2 (0x100-0x107)
                    if 0 <= group_id <= 7:
                        angle_type = None
                        angle_value = None
                        
                        # Buscar si hay información TP2
                        if "TP2" in parts:
                            tp2_index = parts.index("TP2")
                            
                            if len(parts) > tp2_index + 2:
                                angle_type = parts[tp2_index + 1]
                                angle_value = parts[tp2_index + 2]
                        else:
                            # Intentar interpretar basado en el formato esperado del TP2
                            if len(parts) >= 6:  # Al menos tenemos CAN_RX_ID_LEN_BYTE1
                                # Verificar si el primer byte podría ser un tipo de ángulo
                                len_idx = 3
                                first_byte_idx = 4
                                
                                if len_idx < len(parts) and first_byte_idx < len(parts):
                                    try:
                                        first_byte = int(parts[first_byte_idx], 16)
                                        char_value = chr(first_byte)
                                        
                                        if char_value in ['R', 'C', 'O']:
                                            angle_type = char_value
                                            
                                            # Intentar construir el valor del ángulo de los bytes restantes
                                            angle_value = ""
                                            for i in range(first_byte_idx + 1, len(parts)):
                                                try:
                                                    byte_value = int(parts[i], 16)
                                                    if 32 <= byte_value <= 126:  # Rango ASCII imprimible
                                                        angle_value += chr(byte_value)
                                                except ValueError:
                                                    pass
                                    except ValueError:
                                        pass
                        
                        # Si pudimos identificar un tipo y valor de ángulo, actualizar la tabla
                        if angle_type and angle_value and angle_type in ['R', 'C', 'O']:
                            now = datetime.now()
                            
                            # Actualizar el timestamp para este grupo y tipo de ángulo
                            if group_id in self.last_update_times:
                                self.last_update_times[group_id][angle_type] = now
                                self.last_update_times[group_id]['any'] = now
                            
                            # Actualizar el valor en la tabla según el tipo de ángulo
                            item_id = self.tp2_tree.get_children()[group_id]
                            current_values = self.tp2_tree.item(item_id, 'values')
                            new_values = list(current_values)
                            
                            if angle_type == 'R':
                                new_values[1] = angle_value + "°"  # Rolido valor
                                new_values[2] = "Ahora"  # Rolido tiempo
                            elif angle_type == 'C':
                                new_values[3] = angle_value + "°"  # Cabeceo valor
                                new_values[4] = "Ahora"  # Cabeceo tiempo
                            elif angle_type == 'O':
                                new_values[5] = angle_value + "°"  # Orientación valor
                                new_values[6] = "Ahora"  # Orientación tiempo
                            
                            # Actualizar timestamp de última actualización
                            new_values[7] = "Ahora"
                            
                            # Marcar la fila como activa
                            self.tp2_tree.item(item_id, values=tuple(new_values), tags=('active',))
            except Exception as e:
                print(f"Error al procesar mensaje TP2: {str(e)}")
    
    def send_can_message(self):
        """Envía un mensaje CAN usando ID y datos personalizados"""
        if not self.is_connected:
            messagebox.showwarning("No conectado", "Conecte primero al puerto serial")
            return
        
        try:
            can_id = self.can_id_entry.get().strip()
            can_data = self.can_data_entry.get().strip()
            
            # Validar ID hexadecimal
            try:
                id_val = int(can_id, 16)
                if id_val < 0 or id_val > 0x7FF:
                    raise ValueError("ID debe estar entre 0 y 0x7FF")
            except ValueError:
                messagebox.showerror("Error", "ID debe ser un valor hexadecimal válido")
                return
            
            # Validar datos
            if not can_data:
                messagebox.showerror("Error", "Datos no pueden estar vacíos")
                return
            
            # Convertir datos a formato hexadecimal separado por guiones bajos
            data_bytes = []
            for i in range(0, len(can_data), 2):
                if i + 1 < len(can_data):
                    byte_str = can_data[i:i+2]
                    try:
                        byte_val = int(byte_str, 16)
                        data_bytes.append(byte_str)
                    except ValueError:
                        messagebox.showerror("Error", f"Valor de byte inválido: {byte_str}")
                        return
            
            # Formato de comando: SEND_ID_BYTE1_BYTE2_...
            cmd = f"SEND_{can_id}"
            for byte in data_bytes:
                cmd += f"_{byte}"
            
            self.serial_port.write((cmd + "\n").encode('utf-8'))
            # Color verde para mensajes enviados
            self.rx_text.insert(tk.END, f"Enviando: {cmd}\n", "tx_msg")
            self.rx_text.see(tk.END)
            
        except Exception as e:
            messagebox.showerror("Error al enviar", str(e))
    
    def send_tp2_angle(self):
        """Envía un mensaje con formato de ángulo TP2"""
        if not self.is_connected:
            messagebox.showwarning("No conectado", "Conecte primero al puerto serial")
            return
        
        try:
            # Obtener grupo seleccionado para establecer el ID
            group_id = int(self.group_combo.get())
            angle_type = self.angle_type.get()
            angle_value = self.angle_value.get()
            
            try:
                val = int(angle_value)
                if val < -179 or val > 180:
                    raise ValueError("El ángulo debe estar entre -179 y 180")
            except ValueError:
                messagebox.showerror("Error", "Valor de ángulo inválido")
                return
            
            # Formato: SEND_ID_TYPE_VALUES (convertido a hexadecimal)
            can_id = f"{0x100 + group_id:x}"
            
            # Convertir tipo de ángulo y valor a bytes hexadecimales
            data_bytes = []
            
            # Primer byte: tipo de ángulo (R, C, O)
            data_bytes.append(f"{ord(angle_type):02x}")
            
            # Bytes siguientes: valor del ángulo como caracteres ASCII
            for char in angle_value:
                data_bytes.append(f"{ord(char):02x}")
            
            # Construir comando CAN
            cmd = f"SEND_{can_id}"
            for byte in data_bytes:
                cmd += f"_{byte}"
            
            self.serial_port.write((cmd + "\n").encode('utf-8'))
            # Color verde para mensajes enviados
            self.rx_text.insert(tk.END, f"Enviando ángulo TP2 (Grupo {group_id}): {angle_type}={angle_value}°\n", "tx_msg")
            self.rx_text.see(tk.END)
            
        except Exception as e:
            messagebox.showerror("Error al enviar ángulo", str(e))
    
    def set_can_mode(self, mode):
        """Cambia el modo del controlador CAN"""
        if not self.is_connected:
            messagebox.showwarning("No conectado", "Conecte primero al puerto serial")
            return
        
        try:
            cmd = f"MODE_{mode}"
            self.serial_port.write((cmd + "\n").encode('utf-8'))
            # Color verde para mensajes enviados
            self.rx_text.insert(tk.END, f"Cambiando a modo CAN: {mode}\n", "tx_msg")
            self.rx_text.see(tk.END)
        except Exception as e:
            messagebox.showerror("Error al cambiar modo", str(e))
    
    def start_timestamp_updates(self):
        """Inicia la actualización periódica de timestamps en la tabla"""
        self.update_timestamps()
        self.update_timer = self.root.after(1000, self.start_timestamp_updates)
    
    def stop_timestamp_updates(self):
        """Detiene la actualización periódica de timestamps"""
        if self.update_timer:
            self.root.after_cancel(self.update_timer)
            self.update_timer = None
    
    def update_timestamps(self):
        """Actualiza los tiempos mostrados en la tabla desde la última actualización"""
        now = datetime.now()
        
        for group_id in range(8):
            if group_id in self.last_update_times:
                timestamps = self.last_update_times[group_id]
                
                # Obtener el elemento correspondiente en la tabla
                item_id = self.tp2_tree.get_children()[group_id]
                current_values = self.tp2_tree.item(item_id, 'values')
                new_values = list(current_values)
                
                # Variables para controlar el estado visual de cada valor
                r_stale = True
                c_stale = True
                o_stale = True
                all_stale = True
                
                # Actualizar tiempos para cada tipo de ángulo
                for angle_type, timestamp in timestamps.items():
                    if timestamp is None:
                        continue  # No hay actualización registrada
                    
                    # Calcular tiempo transcurrido
                    elapsed = now - timestamp
                    elapsed_seconds = int(elapsed.total_seconds())
                    
                    time_str = ""
                    if elapsed_seconds < 60:
                        time_str = f"{elapsed_seconds}s"
                    elif elapsed_seconds < 3600:
                        time_str = f"{elapsed_seconds // 60}m {elapsed_seconds % 60}s"
                    else:
                        time_str = f"{elapsed_seconds // 3600}h {(elapsed_seconds % 3600) // 60}m"
                    
                    # Determinar si el valor está desactualizado (más de 2 segundos)
                    is_stale = elapsed_seconds > 2
                    
                    # Actualizar el campo correspondiente
                    if angle_type == 'R':
                        new_values[2] = time_str
                        r_stale = is_stale
                    elif angle_type == 'C':
                        new_values[4] = time_str
                        c_stale = is_stale
                    elif angle_type == 'O':
                        new_values[6] = time_str
                        o_stale = is_stale
                    elif angle_type == 'any':
                        new_values[7] = time_str
                        all_stale = is_stale
                
                # Actualizar los valores en la tabla
                self.tp2_tree.item(item_id, values=tuple(new_values))
                
                # Aplicar el tag correspondiente a la fila según el estado de actualización
                if all_stale:
                    self.tp2_tree.item(item_id, tags=('stale',))
                else:
                    self.tp2_tree.item(item_id, tags=('active',))
    
    def reset_tp2_data(self):
        """Restablece todos los datos de TP2 a su estado inicial"""
        for i in range(8):
            item_id = self.tp2_tree.get_children()[i]
            self.tp2_tree.item(item_id, values=(i, '--', 'Nunca', '--', 'Nunca', '--', 'Nunca', 'Nunca'), tags=('stale',))
            self.last_update_times[i] = {
                'R': None,
                'C': None,
                'O': None,
                'any': None
            }

if __name__ == "__main__":
    root = tk.Tk()
    app = CanMonitorApp(root)
    root.mainloop()