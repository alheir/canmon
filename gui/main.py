import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import serial
import serial.tools.list_ports
import threading
import time
import re
import platform
import os
import random
import math
from datetime import datetime
from collections import deque
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.animation as animation
import numpy as np

class PlotWindow:
    def __init__(self, parent, group_id, data_source):
        self.window = tk.Toplevel(parent)
        self.window.title(f"Group {group_id} - Angle Data Plot")
        self.window.geometry("900x700")
        self.group_id = group_id
        self.data_source = data_source
        
        # How many seconds of data to show
        self.time_window = 30
        
        # Setup figures for plotting
        self.setup_plots()
        
        # Start animation
        self.ani = animation.FuncAnimation(
            self.fig, self.update_plots, interval=100,  # Update every 100ms
            blit=False)
        
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
    def setup_plots(self):
        self.fig = Figure(figsize=(9, 7), dpi=100)
        
        # Create three subplots for roll, pitch, and orientation
        self.roll_plot = self.fig.add_subplot(311)
        self.roll_plot.set_title("Roll")
        self.roll_plot.set_ylabel("Degrees")
        self.roll_plot.grid(True)
        
        self.pitch_plot = self.fig.add_subplot(312)
        self.pitch_plot.set_title("Pitch")
        self.pitch_plot.set_ylabel("Degrees")
        self.pitch_plot.grid(True)
        
        self.orient_plot = self.fig.add_subplot(313)
        self.orient_plot.set_title("Orientation")
        self.orient_plot.set_ylabel("Degrees")
        self.orient_plot.set_xlabel("Time (seconds)")
        self.orient_plot.grid(True)
        
        # Initial empty lines
        self.roll_line, = self.roll_plot.plot([], [], 'b-')
        self.pitch_line, = self.pitch_plot.plot([], [], 'g-')
        self.orient_line, = self.orient_plot.plot([], [], 'r-')
        
        # Common y-axis range for all plots
        for plot in [self.roll_plot, self.pitch_plot, self.orient_plot]:
            plot.set_ylim(-180, 180)
            plot.axhline(y=0, color='k', linestyle='--', alpha=0.3)
        
        self.fig.tight_layout()
        
        # Add the plot to the window
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.window)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # For toggling plot visibility
        self.visible_plots = {'roll': True, 'pitch': True, 'orientation': True}
        
        # Add controls for toggling plots
        self.control_frame = ttk.Frame(self.window)
        self.control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Checkbuttons to toggle plot visibility
        self.roll_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.control_frame, text="Show Roll", 
                       variable=self.roll_var, 
                       command=lambda: self.toggle_plot('roll')).pack(side=tk.LEFT, padx=5)
        
        self.pitch_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.control_frame, text="Show Pitch", 
                        variable=self.pitch_var, 
                        command=lambda: self.toggle_plot('pitch')).pack(side=tk.LEFT, padx=5)
        
        self.orient_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.control_frame, text="Show Orientation", 
                        variable=self.orient_var, 
                        command=lambda: self.toggle_plot('orientation')).pack(side=tk.LEFT, padx=5)
        
    def toggle_plot(self, plot_name):
        if plot_name == 'roll':
            self.visible_plots['roll'] = self.roll_var.get()
        elif plot_name == 'pitch':
            self.visible_plots['pitch'] = self.pitch_var.get()
        elif plot_name == 'orientation':
            self.visible_plots['orientation'] = self.orient_var.get()
    
    def update_plots(self, frame):
        # Get data for this group
        data = self.data_source.get_plot_data(self.group_id)
        
        # Current time for reference
        now = time.time()
        
        # Filter data to show only the last time_window seconds
        roll_data = [(t, v) for t, v in data['R'] if now - t <= self.time_window]
        pitch_data = [(t, v) for t, v in data['C'] if now - t <= self.time_window]
        orient_data = [(t, v) for t, v in data['O'] if now - t <= self.time_window]
        
        # Update each plot
        if roll_data and self.visible_plots['roll']:
            times, values = zip(*roll_data) if roll_data else ([], [])
            # Convert absolute timestamps to seconds ago
            rel_times = [now - t for t in times]
            self.roll_line.set_data(rel_times, values)
            self.roll_plot.set_xlim(self.time_window, 0)  # Reversed x-axis (newest data on right)
        else:
            self.roll_line.set_data([], [])
            
        if pitch_data and self.visible_plots['pitch']:
            times, values = zip(*pitch_data) if pitch_data else ([], [])
            rel_times = [now - t for t in times]
            self.pitch_line.set_data(rel_times, values)
            self.pitch_plot.set_xlim(self.time_window, 0)
        else:
            self.pitch_line.set_data([], [])
            
        if orient_data and self.visible_plots['orientation']:
            times, values = zip(*orient_data) if orient_data else ([], [])
            rel_times = [now - t for t in times]
            self.orient_line.set_data(rel_times, values)
            self.orient_plot.set_xlim(self.time_window, 0)
        else:
            self.orient_line.set_data([], [])
        
        return self.roll_line, self.pitch_line, self.orient_line
    
    def on_close(self):
        # Stop the animation when window is closed
        self.ani.event_source.stop()
        self.window.destroy()

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
        self.port_info = {}  # Stores detailed port information
        self.last_update_times = {}  # Stores timestamps of updates
        self.update_timer = None  # For periodic timestamp updates
        
        # Historical data for plotting
        self.plot_data = {}
        for i in range(8):
            self.plot_data[i] = {
                'R': deque(maxlen=500),  # Store up to 500 points
                'C': deque(maxlen=500),
                'O': deque(maxlen=500)
            }
        
        # Reference to plot windows
        self.plot_windows = {}
        
        # Variables for random transmission
        self.random_transmission_active = False
        self.random_transmission_thread = None
        self.random_last_values = {'R': 0, 'C': 0, 'O': 0}
        self.random_last_sent_time = {'R': 0, 'C': 0, 'O': 0}
        
        # Create interface
        self.create_widgets()
        
        # Update COM port list
        self.refresh_ports()
    
    def get_plot_data(self, group_id):
        """Retrieves plot data for a specific group (used by PlotWindow)"""
        if group_id in self.plot_data:
            return self.plot_data[group_id]
        return {'R': deque(), 'C': deque(), 'O': deque()}

    def create_widgets(self):
        # Main frame with two columns
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === LEFT COLUMN ===
        left_frame = ttk.LabelFrame(main_frame, text="Connection and Control", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Serial connection section
        conn_frame = ttk.Frame(left_frame)
        conn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(conn_frame, text="Port:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        self.port_combo = ttk.Combobox(conn_frame, width=25)
        self.port_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.port_combo.bind('<<ComboboxSelected>>', self.on_port_selected)
        
        self.refresh_btn = ttk.Button(conn_frame, text="Refresh", command=self.refresh_ports)
        self.refresh_btn.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        
        self.connect_btn = ttk.Button(conn_frame, text="Connect", command=self.toggle_connection)
        self.connect_btn.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        # Port information
        self.port_info_label = ttk.Label(conn_frame, text="", wraplength=300)
        self.port_info_label.grid(row=1, column=0, columnspan=4, sticky=tk.W, padx=5, pady=5)
        
        # Section for sending custom CAN messages
        send_frame = ttk.LabelFrame(left_frame, text="Send CAN Message", padding=10)
        send_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(send_frame, text="ID (hex):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.can_id_entry = ttk.Entry(send_frame, width=5)
        self.can_id_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.can_id_entry.insert(0, "100")
        
        ttk.Label(send_frame, text="Data (hex):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.can_data_entry = ttk.Entry(send_frame, width=20)
        self.can_data_entry.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        self.send_btn = ttk.Button(send_frame, text="Send", command=self.send_can_message)
        self.send_btn.grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)
        
        # Section for TP2 presets
        presets_frame = ttk.LabelFrame(left_frame, text="TP2 Presets", padding=10)
        presets_frame.pack(fill=tk.X, pady=10)
        
        # Group selection
        ttk.Label(presets_frame, text="Group:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.group_combo = ttk.Combobox(presets_frame, width=5, values=[f"{i}" for i in range(8)])
        self.group_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.group_combo.current(0)
        self.group_combo.bind("<<ComboboxSelected>>", self.on_group_selected)
        
        # Input method selection
        self.input_method = tk.StringVar(value="numeric")
        ttk.Label(presets_frame, text="Input Method:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Radiobutton(presets_frame, text="Numeric", variable=self.input_method, 
                        value="numeric", command=self.toggle_input_method).grid(row=1, column=1)
        ttk.Radiobutton(presets_frame, text="String Format", variable=self.input_method, 
                        value="string", command=self.toggle_input_method).grid(row=1, column=2, columnspan=2)
        
        # === Numeric Input (existing) ===
        self.numeric_frame = ttk.Frame(presets_frame)
        self.numeric_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        # Angle type
        ttk.Label(self.numeric_frame, text="Type:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.angle_type = tk.StringVar(value="R")
        ttk.Radiobutton(self.numeric_frame, text="Roll (R)", variable=self.angle_type, value="R").grid(row=0, column=1)
        ttk.Radiobutton(self.numeric_frame, text="Pitch (C)", variable=self.angle_type, value="C").grid(row=0, column=2)
        ttk.Radiobutton(self.numeric_frame, text="Orientation (O)", variable=self.angle_type, value="O").grid(row=0, column=3)
        
        # Angle value
        ttk.Label(self.numeric_frame, text="Angle:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.angle_value = ttk.Spinbox(self.numeric_frame, from_=-179, to=180, width=5)
        self.angle_value.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        self.angle_value.set("0")
        
        # === String Format Input (new) ===
        self.string_frame = ttk.Frame(presets_frame)
        self.string_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        ttk.Label(self.string_frame, text="Angle String:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.angle_string = ttk.Entry(self.string_frame, width=10)
        self.angle_string.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(self.string_frame, 
                 text="Format: R|C|O followed by angle value\nExamples: R-34, C0, O67, R+138").grid(
                 row=0, column=2, rowspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Send button (shared)
        self.send_preset_btn = ttk.Button(presets_frame, text="Send Angle", command=self.send_tp2_angle)
        self.send_preset_btn.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Initially hide the string input frame
        self.string_frame.grid_remove()
        
        # CAN Mode
        mode_frame = ttk.LabelFrame(left_frame, text="CAN Mode", padding=10)
        mode_frame.pack(fill=tk.X, pady=10)
        
        self.normal_mode_btn = ttk.Button(mode_frame, text="Normal Mode", command=lambda: self.set_can_mode("NORMAL"))
        self.normal_mode_btn.grid(row=0, column=0, padx=5, pady=5)
        
        self.loopback_mode_btn = ttk.Button(mode_frame, text="Loopback Mode", command=lambda: self.set_can_mode("LOOPBACK"))
        self.loopback_mode_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # Add Random Transmission section to left frame
        random_frame = ttk.LabelFrame(left_frame, text="Random Transmission (TP2 Timing)", padding=10)
        random_frame.pack(fill=tk.X, pady=10)
        
        # Group selection for random transmission
        ttk.Label(random_frame, text="Source Group:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.random_group_combo = ttk.Combobox(random_frame, width=5, values=[f"{i}" for i in range(8)])
        self.random_group_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.random_group_combo.current(0)
        
        # Frame for displaying current random values
        values_frame = ttk.Frame(random_frame)
        values_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(values_frame, text="Current Random Values:").grid(row=0, column=0, sticky=tk.W)
        
        # Labels to show current random values
        self.random_roll_val = ttk.Label(values_frame, text="Roll: 0°")
        self.random_roll_val.grid(row=1, column=0, sticky=tk.W, padx=5)
        
        self.random_pitch_val = ttk.Label(values_frame, text="Pitch: 0°")
        self.random_pitch_val.grid(row=1, column=1, sticky=tk.W, padx=5)
        
        self.random_orient_val = ttk.Label(values_frame, text="Orient: 0°")
        self.random_orient_val.grid(row=1, column=2, sticky=tk.W, padx=5)
        
        # Status label for timing info
        self.random_status = ttk.Label(random_frame, text="Idle", width=50)
        self.random_status.grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
        
        # Start/Stop button
        self.random_btn = ttk.Button(random_frame, text="Start Random Transmission", 
                                    command=self.toggle_random_transmission)
        self.random_btn.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Add a separator
        ttk.Separator(random_frame, orient=tk.HORIZONTAL).grid(
            row=4, column=0, columnspan=3, sticky=tk.EW, pady=10)
        
        # Information about timing rules
        timing_text = ("Timing rules:\n"
                      "• Max 20 packets/second\n"
                      "• Send immediately if angle changes ≥5°\n"
                      "• Send at least every 2 seconds\n"
                      "• Roll, pitch and orientation treated independently")
        ttk.Label(random_frame, text=timing_text, justify=tk.LEFT).grid(
            row=5, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
        
        # === RIGHT COLUMN ===
        right_frame = ttk.LabelFrame(main_frame, text="CAN Messages", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Area to display received messages
        self.rx_text = scrolledtext.ScrolledText(right_frame, width=50, height=20)
        self.rx_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Configure colors for messages
        self.rx_text.tag_config("tx_msg", foreground="green")
        self.rx_text.tag_config("rx_msg", foreground="blue")
        self.rx_text.tag_config("system", foreground="black")
        self.rx_text.tag_config("error", foreground="red")
        
        # Buttons to clear and enable/disable autoscroll
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.clear_btn = ttk.Button(btn_frame, text="Clear", command=lambda: self.rx_text.delete(1.0, tk.END))
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Area for interpreted TP2 messages
        tp2_frame = ttk.LabelFrame(right_frame, text="Interpreted TP2 Messages", padding=10)
        tp2_frame.pack(fill=tk.X, pady=10)
        
        # Table to display angles
        columns = ('group', 'roll', 'roll_time', 'pitch', 'pitch_time', 'orientation', 'orientation_time', 'last_update')
        self.tp2_tree = ttk.Treeview(tp2_frame, columns=columns, show='headings', height=8)
        
        # Define headers
        self.tp2_tree.heading('group', text='Group')
        self.tp2_tree.heading('roll', text='Roll')
        self.tp2_tree.heading('roll_time', text='Roll Time')
        self.tp2_tree.heading('pitch', text='Pitch')
        self.tp2_tree.heading('pitch_time', text='Pitch Time')
        self.tp2_tree.heading('orientation', text='Orientation')
        self.tp2_tree.heading('orientation_time', text='Orientation Time')
        self.tp2_tree.heading('last_update', text='Last Update')
        
        # Adjust column widths
        self.tp2_tree.column('group', width=50, anchor=tk.CENTER)
        self.tp2_tree.column('roll', width=60, anchor=tk.CENTER)
        self.tp2_tree.column('roll_time', width=80, anchor=tk.CENTER)
        self.tp2_tree.column('pitch', width=60, anchor=tk.CENTER)
        self.tp2_tree.column('pitch_time', width=80, anchor=tk.CENTER)
        self.tp2_tree.column('orientation', width=60, anchor=tk.CENTER)
        self.tp2_tree.column('orientation_time', width=80, anchor=tk.CENTER)
        self.tp2_tree.column('last_update', width=100, anchor=tk.CENTER)
        
        # Styles for the table
        self.tp2_tree.tag_configure('stale', foreground='gray')
        self.tp2_tree.tag_configure('active', foreground='black')
        
        self.tp2_tree.pack(fill=tk.BOTH, expand=True)
        
        # Initialize with groups 0 to 7 (according to TP2)
        for i in range(8):
            self.tp2_tree.insert('', tk.END, values=(i, '--', 'Never', '--', 'Never', '--', 'Never', 'Never'), tags=('stale',))
            self.last_update_times[i] = {
                'R': None,
                'C': None,
                'O': None,
                'any': None
            }
        
        # Button to open plotting window in the TP2 section
        plotting_frame = ttk.LabelFrame(right_frame, text="Real-time Plotting", padding=10)
        plotting_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(plotting_frame, text="Select Group:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.plot_group_combo = ttk.Combobox(plotting_frame, width=5, values=[f"{i}" for i in range(8)])
        self.plot_group_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.plot_group_combo.current(0)
        
        self.open_plot_btn = ttk.Button(plotting_frame, text="Open Plot Window", 
                                        command=self.open_plot_window)
        self.open_plot_btn.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

    def toggle_input_method(self):
        """Toggles between numeric and string input methods"""
        if self.input_method.get() == "numeric":
            self.string_frame.grid_remove()
            self.numeric_frame.grid()
        else:
            self.numeric_frame.grid_remove()
            self.string_frame.grid()
    
    def toggle_random_transmission(self):
        """Starts or stops the random transmission mode"""
        if not self.is_connected:
            messagebox.showwarning("Not Connected", "Connect to the serial port first")
            return
        
        if not self.random_transmission_active:
            # Start random transmission
            group_id = int(self.random_group_combo.get())
            self.random_transmission_active = True
            self.random_btn.config(text="Stop Random Transmission")
            
            # Reset values
            self.random_last_values = {'R': 0, 'C': 0, 'O': 0}
            self.random_last_sent_time = {'R': 0, 'C': 0, 'O': 0}
            
            # Initialize sine wave parameters (different for each angle type)
            self.sine_params = {
                'R': {
                    'amplitude': random.randint(20, 60),  # Range of motion
                    'period': random.uniform(5, 15),      # Seconds per cycle
                    'phase': random.uniform(0, 2*math.pi), # Starting phase
                    'offset': random.randint(-30, 30)     # Center offset
                },
                'C': {
                    'amplitude': random.randint(15, 45),
                    'period': random.uniform(3, 10),
                    'phase': random.uniform(0, 2*math.pi),
                    'offset': random.randint(-20, 20)
                },
                'O': {
                    'amplitude': random.randint(50, 120),
                    'period': random.uniform(8, 20),
                    'phase': random.uniform(0, 2*math.pi),
                    'offset': random.randint(-50, 50)
                }
            }
            
            # Start the thread
            self.random_transmission_thread = threading.Thread(
                target=self.random_transmission_loop,
                args=(group_id,),
                daemon=True
            )
            self.random_transmission_thread.start()
            
            # Log
            self.rx_text.insert(tk.END, f"Started sinusoidal transmission for Group {group_id}\n", "system")
            self.rx_text.see(tk.END)
        else:
            # Stop random transmission
            self.random_transmission_active = False
            self.random_btn.config(text="Start Random Transmission")
            self.random_status.config(text="Idle")
            
            # Log
            self.rx_text.insert(tk.END, "Stopped sinusoidal transmission\n", "system")
            self.rx_text.see(tk.END)
    
    def random_transmission_loop(self, group_id):
        """Thread function to send sinusoidal angle values with TP2 timing rules"""
        last_global_send = time.time()  # For minimum interval between packets (max 20/sec)
        start_time = time.time()        # Reference time for sine waves
        next_angle_index = 0            # To ensure round-robin between angle types
        
        angle_types = ['R', 'C', 'O']
        
        while self.random_transmission_active and self.is_connected:
            current_time = time.time()
            elapsed = current_time - start_time
            
            # Determine the next angle type to check (round-robin)
            angle_type = angle_types[next_angle_index]
            next_angle_index = (next_angle_index + 1) % 3
            
            # Calculate new value based on sine wave with some randomness
            params = self.sine_params[angle_type]
            sine_value = params['amplitude'] * math.sin(2 * math.pi * elapsed / params['period'] + params['phase']) + params['offset']
            
            # Add small random variation to make it more natural (±2 degrees)
            new_value = int(sine_value + random.uniform(-2, 2))
            
            # Ensure value stays within range
            new_value = max(-179, min(180, new_value))
            
            # Update display for this angle type
            if angle_type == 'R':
                self.root.after(0, lambda v=new_value: self.random_roll_val.config(text=f"Roll: {v}°"))
            elif angle_type == 'C':
                self.root.after(0, lambda v=new_value: self.random_pitch_val.config(text=f"Pitch: {v}°"))
            else: # 'O'
                self.root.after(0, lambda v=new_value: self.random_orient_val.config(text=f"Orient: {v}°"))
            
            # Time since last send for this angle
            time_since_last = current_time - self.random_last_sent_time.get(angle_type, 0)
            
            # Check if we need to send this angle
            should_send = False
            reason = ""
            
            # Rule 1: If the angle changed by 5 degrees or more
            if abs(new_value - self.random_last_values.get(angle_type, 0)) >= 5:
                should_send = True
                reason = "≥5° change"
            
            # Rule 2: If we haven't sent this angle in 2 seconds
            elif time_since_last >= 2.0:
                should_send = True
                reason = "2s timeout"
            
            # Only send if minimum interval passed (max 20 packets/second)
            if should_send and (current_time - last_global_send) >= 0.250:  # 250ms minimum interval (aiming for 50ms)
                # Send the angle
                can_id = f"{0x100 + group_id:x}"
                
                # Format as string (angleType + value)
                angle_string = f"{angle_type}{new_value}"
                
                # Convert to bytes
                data_bytes = []
                for char in angle_string:
                    data_bytes.append(f"{ord(char):02x}")
                
                # Build command
                cmd = f"SEND_{can_id}"
                for byte in data_bytes:
                    cmd += f"_{byte}"
                
                # Send the command
                try:
                    self.serial_port.write((cmd + "\n").encode('utf-8'))
                    
                    # Update status with information about this packet
                    status_text = f"Sent {angle_type}={new_value}° ({reason})"
                    self.root.after(0, lambda t=status_text: self.random_status.config(text=t))
                    
                    # Update sent time and value
                    self.random_last_sent_time[angle_type] = current_time
                    self.random_last_values[angle_type] = new_value
                    last_global_send = current_time
                    
                    # Log to message window
                    msg = f"Sine: Sent {angle_type}={new_value}° for Group {group_id} ({reason})"
                    self.root.after(0, lambda m=msg: self.rx_text.insert(tk.END, f"{m}\n", "tx_msg"))
                    self.root.after(0, self.rx_text.see, tk.END)
                    
                    # Wait a bit before trying the next angle (respect max 20 packets/sec)
                    time.sleep(0.05)
                except Exception as e:
                    # Error handling
                    self.root.after(0, lambda: self.rx_text.insert(
                        tk.END, f"Error sending angle: {str(e)}\n", "error"))
                    self.root.after(0, self.rx_text.see, tk.END)
                    
                    # May need to stop if serial connection is lost
                    if not self.is_connected:
                        self.random_transmission_active = False
                        break
            
            # Small sleep to prevent CPU overuse
            time.sleep(0.02)
    
    def on_closing(self):
        """Cleanup when the application is closing"""
        # Stop random transmission if active
        self.random_transmission_active = False
        
        # Disconnect if connected
        if self.is_connected:
            self.toggle_connection()
        
        # Close plot windows
        for window in self.plot_windows.values():
            if hasattr(window, 'window') and window.window.winfo_exists():
                window.window.destroy()
        
        # Close main window
        self.root.destroy()

    def on_port_selected(self, event):
        """Displays detailed information about the selected port"""
        selected = self.port_combo.get()
        if selected in self.port_info:
            info = self.port_info[selected]
            info_text = f"Port: {info['device']}\n"
            if info['description']:
                info_text += f"Description: {info['description']}\n"
            if info['manufacturer']:
                info_text += f"Manufacturer: {info['manufacturer']}\n"
            if info['hwid']:
                info_text += f"Hardware ID: {info['hwid']}\n"
            if info['serial_number'] and info['serial_number'] != 'None':
                info_text += f"Serial Number: {info['serial_number']}\n"
            
            self.port_info_label.config(text=info_text)
        else:
            self.port_info_label.config(text="")
    
    def on_group_selected(self, event):
        """Updates the CAN ID entry when a group is selected"""
        group_id = int(self.group_combo.get())
        can_id = f"{0x100 + group_id:x}"
        self.can_id_entry.delete(0, tk.END)
        self.can_id_entry.insert(0, can_id)
    
    def refresh_ports(self):
        """Updates the list of available serial ports with detailed information"""
        self.port_info = {}
        ports = []
        display_names = []
        
        try:
            for port in serial.tools.list_ports.comports():
                # Create a unique identifier for the port
                port_id = port.device
                
                # Save detailed port information
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
                
                # Create a descriptive name to display in the ComboBox
                display_name = port.device
                if port.description and port.description != port.device:
                    display_name = f"{port.device} - {port.description}"
                
                display_names.append(display_name)
                ports.append(port_id)
        
        except Exception as e:
            messagebox.showerror("Error", f"Error detecting ports: {str(e)}")
        
        # Update the ComboBox
        self.port_combo['values'] = display_names
        if display_names:
            self.port_combo.current(0)
            self.on_port_selected(None)  # Show info for the first port
        else:
            self.port_info_label.config(text="No serial ports detected")
    
    def toggle_connection(self):
        """Connects or disconnects from the serial port"""
        if not self.is_connected:
            selected = self.port_combo.get()
            
            # Extract the device name from the displayed text (may contain description)
            device = selected.split(' - ')[0] if ' - ' in selected else selected
            
            # If it's in the info dictionary, use that port
            if device in self.port_info:
                port = self.port_info[device]['device']
            else:
                # If not in the dictionary, use the selected directly
                port = device
                
            try:
                self.serial_port = serial.Serial(port, 115200, timeout=1)
                self.is_connected = True
                self.connect_btn['text'] = "Disconnect"
                self.should_read = True
                
                # Reset TP2 data on connect
                self.reset_tp2_data()
                
                # Start thread for continuous reading
                self.reading_thread = threading.Thread(target=self.read_serial_data)
                self.reading_thread.daemon = True
                self.reading_thread.start()
                
                # Start periodic timestamp updates
                self.start_timestamp_updates()
                
                # Display system information
                os_info = platform.platform()
                self.rx_text.insert(tk.END, f"System: {os_info}\n", "system")
                self.rx_text.insert(tk.END, f"Connected to {port} @ 115200 bps\n", "system")
                self.rx_text.see(tk.END)
            except Exception as e:
                messagebox.showerror("Connection Error", str(e))
        else:
            self.should_read = False
            if self.serial_port:
                self.serial_port.close()
            self.is_connected = False
            self.connect_btn['text'] = "Connect"
            self.rx_text.insert(tk.END, "Disconnected\n", "system")
            self.rx_text.see(tk.END)
            
            # Reset TP2 data on disconnect
            self.reset_tp2_data()
            
            # Stop timestamp updates
            self.stop_timestamp_updates()
    
    def read_serial_data(self):
        """Reads data from the serial port continuously"""
        while self.should_read:
            if self.serial_port and self.serial_port.in_waiting:
                try:
                    data = self.serial_port.readline().decode('utf-8').strip()
                    self.process_received_data(data)
                except Exception as e:
                    self.root.after(0, lambda: self.rx_text.insert(tk.END, f"Read Error: {str(e)}\n", "error"))
                    self.root.after(0, lambda: self.rx_text.see(tk.END))
            time.sleep(0.01)
    
    def process_received_data(self, data):
        """Processes data received via serial"""
        if not data:
            return
        
        # Add to text area with blue color for reception
        self.root.after(0, lambda: self.rx_text.insert(tk.END, f"{data}\n", "rx_msg"))
        self.root.after(0, lambda: self.rx_text.see(tk.END))
        
        # Check if it's a CAN message in TP2 format
        if data.startswith("CAN_RX_"):
            try:
                # Expected format: CAN_RX_ID_LEN_BYTE1_BYTE2_..._TP2_TYPE_VALUE
                parts = data.split("_")
                
                if len(parts) >= 5:
                    # Extract ID to determine the group
                    id_hex = parts[2]
                    try:
                        group_id = int(id_hex, 16) - 0x100
                    except ValueError:
                        print(f"Error converting ID: {id_hex}")
                        return
                        
                    # Verify if it's an ID within the TP2 range (0x100-0x107)
                    if 0 <= group_id <= 7:
                        angle_type = None
                        angle_value = None
                        
                        # Look for TP2 information
                        if "TP2" in parts:
                            tp2_index = parts.index("TP2")
                            
                            if len(parts) > tp2_index + 2:
                                angle_type = parts[tp2_index + 1]
                                angle_value = parts[tp2_index + 2]
                        else:
                            # Attempt to interpret based on the expected TP2 format
                            if len(parts) >= 6:  # At least we have CAN_RX_ID_LEN_BYTE1
                                # Verify if the first byte could be an angle type
                                len_idx = 3
                                first_byte_idx = 4
                                
                                if len_idx < len(parts) and first_byte_idx < len(parts):
                                    try:
                                        first_byte = int(parts[first_byte_idx], 16)
                                        char_value = chr(first_byte)
                                        
                                        if char_value in ['R', 'C', 'O']:
                                            angle_type = char_value
                                            
                                            # Attempt to construct the angle value from the remaining bytes
                                            angle_value = ""
                                            for i in range(first_byte_idx + 1, len(parts)):
                                                try:
                                                    byte_value = int(parts[i], 16)
                                                    if 32 <= byte_value <= 126:  # Printable ASCII range
                                                        angle_value += chr(byte_value)
                                                except ValueError:
                                                    pass
                                    except ValueError:
                                        pass
                        
                        # If we could identify an angle type and value, update the table
                        if angle_type and angle_value and angle_type in ['R', 'C', 'O']:
                            now = datetime.now()
                            current_time = time.time()
                            
                            # Update the timestamp for this group and angle type
                            if group_id in self.last_update_times:
                                self.last_update_times[group_id][angle_type] = now
                                self.last_update_times[group_id]['any'] = now
                            
                            # Store data for plotting
                            try:
                                # Convert angle value to float and store with timestamp
                                angle_float = float(angle_value)
                                if group_id in self.plot_data and angle_type in self.plot_data[group_id]:
                                    self.plot_data[group_id][angle_type].append((current_time, angle_float))
                            except ValueError:
                                # If conversion fails, don't store for plotting
                                pass
                            
                            # Update the value in the table based on the angle type
                            item_id = self.tp2_tree.get_children()[group_id]
                            current_values = self.tp2_tree.item(item_id, 'values')
                            new_values = list(current_values)
                            
                            if angle_type == 'R':
                                new_values[1] = angle_value + "°"  # Roll value
                                new_values[2] = "Now"  # Roll time
                            elif angle_type == 'C':
                                new_values[3] = angle_value + "°"  # Pitch value
                                new_values[4] = "Now"  # Pitch time
                            elif angle_type == 'O':
                                new_values[5] = angle_value + "°"  # Orientation value
                                new_values[6] = "Now"  # Orientation time
                            
                            # Update last update timestamp
                            new_values[7] = "Now"
                            
                            # Mark the row as active
                            self.tp2_tree.item(item_id, values=tuple(new_values), tags=('active',))
            except Exception as e:
                print(f"Error processing TP2 message: {str(e)}")
    
    def send_can_message(self):
        """Sends a CAN message using custom ID and data"""
        if not self.is_connected:
            messagebox.showwarning("Not Connected", "Connect to the serial port first")
            return
        
        try:
            can_id = self.can_id_entry.get().strip()
            can_data = self.can_data_entry.get().strip()
            
            # Validate hexadecimal ID
            try:
                id_val = int(can_id, 16)
                if id_val < 0 or id_val > 0x7FF:
                    raise ValueError("ID must be between 0 and 0x7FF")
            except ValueError:
                messagebox.showerror("Error", "ID must be a valid hexadecimal value")
                return
            
            # Validate data
            if not can_data:
                messagebox.showerror("Error", "Data cannot be empty")
                return
            
            # Convert data to hexadecimal format separated by underscores
            data_bytes = []
            for i in range(0, len(can_data), 2):
                if i + 1 < len(can_data):
                    byte_str = can_data[i:i+2]
                    try:
                        byte_val = int(byte_str, 16)
                        data_bytes.append(byte_str)
                    except ValueError:
                        messagebox.showerror("Error", f"Invalid byte value: {byte_str}")
                        return
            
            # Command format: SEND_ID_BYTE1_BYTE2_...
            cmd = f"SEND_{can_id}"
            for byte in data_bytes:
                cmd += f"_{byte}"
            
            self.serial_port.write((cmd + "\n").encode('utf-8'))
            # Green color for sent messages
            self.rx_text.insert(tk.END, f"Sending: {cmd}\n", "tx_msg")
            self.rx_text.see(tk.END)
            
        except Exception as e:
            messagebox.showerror("Error Sending", str(e))
    
    def send_tp2_angle(self):
        """Sends a message with TP2 angle format"""
        if not self.is_connected:
            messagebox.showwarning("Not Connected", "Connect to the serial port first")
            return
        
        try:
            # Get selected group to set the ID
            group_id = int(self.group_combo.get())
            can_id = f"{0x100 + group_id:x}"
            
            # Handle based on selected input method
            if self.input_method.get() == "numeric":
                # Original numerical input handling
                angle_type = self.angle_type.get()
                angle_value = self.angle_value.get()
                
                try:
                    val = int(angle_value)
                    if val < -179 or val > 180:
                        raise ValueError("Angle must be between -179 and 180")
                except ValueError:
                    messagebox.showerror("Error", "Invalid angle value")
                    return
                
                # Convert angle type and value to hexadecimal bytes
                data_bytes = []
                
                # First byte: angle type (R, C, O)
                data_bytes.append(f"{ord(angle_type):02x}")
                
                # Next bytes: angle value as ASCII characters
                for char in angle_value:
                    data_bytes.append(f"{ord(char):02x}")
                
                display_msg = f"Sending TP2 angle (Group {group_id}): {angle_type}={angle_value}°"
                
            else:
                # New string format input handling
                angle_string = self.angle_string.get().strip()
                
                # Validate the angle string
                if not self.validate_angle_string(angle_string):
                    messagebox.showerror("Error", "Invalid angle string format.\n"
                                         "Format should be R|C|O followed by 1-4 digits.\n"
                                         "Examples: R-34, C0, O67, R+138")
                    return
                
                # Convert the entire string to hex bytes
                data_bytes = []
                for char in angle_string:
                    data_bytes.append(f"{ord(char):02x}")
                
                display_msg = f"Sending TP2 angle string (Group {group_id}): {angle_string}"
            
            # Build CAN command
            cmd = f"SEND_{can_id}"
            for byte in data_bytes:
                cmd += f"_{byte}"
            
            self.serial_port.write((cmd + "\n").encode('utf-8'))
            # Green color for sent messages
            self.rx_text.insert(tk.END, f"{display_msg}\n", "tx_msg")
            self.rx_text.see(tk.END)
            
        except Exception as e:
            messagebox.showerror("Error Sending Angle", str(e))
    
    def validate_angle_string(self, angle_string):
        """Validates if the provided string follows the angle string format"""
        if not angle_string or len(angle_string) < 2 or len(angle_string) > 5:
            return False
        
        # First character must be one of R, C, or O
        if angle_string[0] not in ['R', 'C', 'O']:
            return False
        
        # Rest of the string must form a valid angle value (1-4 characters)
        value_part = angle_string[1:]
        
        # Check if it starts with a sign (optional)
        if value_part.startswith('+') or value_part.startswith('-'):
            value_part = value_part[1:]  # Remove the sign
            if not value_part:  # If there's nothing after the sign
                return False
        
        # Remaining characters should be digits
        if not value_part.isdigit():
            return False
        
        # Valid string format
        return True
    
    def set_can_mode(self, mode):
        """Changes the CAN controller mode"""
        if not self.is_connected:
            messagebox.showwarning("Not Connected", "Connect to the serial port first")
            return
        
        try:
            cmd = f"MODE_{mode}"
            self.serial_port.write((cmd + "\n").encode('utf-8'))
            # Green color for sent messages
            self.rx_text.insert(tk.END, f"Changing CAN mode: {mode}\n", "tx_msg")
            self.rx_text.see(tk.END)
        except Exception as e:
            messagebox.showerror("Error Changing Mode", str(e))
    
    def start_timestamp_updates(self):
        """Starts periodic timestamp updates in the table"""
        self.update_timestamps()
        self.update_timer = self.root.after(1000, self.start_timestamp_updates)
    
    def stop_timestamp_updates(self):
        """Stops periodic timestamp updates"""
        if self.update_timer:
            self.root.after_cancel(self.update_timer)
            self.update_timer = None
    
    def update_timestamps(self):
        """Updates the times displayed in the table since the last update"""
        now = datetime.now()
        
        for group_id in range(8):
            if group_id in self.last_update_times:
                timestamps = self.last_update_times[group_id]
                
                # Get the corresponding item in the table
                item_id = self.tp2_tree.get_children()[group_id]
                current_values = self.tp2_tree.item(item_id, 'values')
                new_values = list(current_values)
                
                # Variables to control the visual state of each value
                r_stale = True
                c_stale = True
                o_stale = True
                all_stale = True
                
                # Update times for each angle type
                for angle_type, timestamp in timestamps.items():
                    if timestamp is None:
                        continue  # No update recorded
                    
                    # Calculate elapsed time
                    elapsed = now - timestamp
                    elapsed_seconds = int(elapsed.total_seconds())
                    
                    time_str = ""
                    if elapsed_seconds < 60:
                        time_str = f"{elapsed_seconds}s"
                    elif elapsed_seconds < 3600:
                        time_str = f"{elapsed_seconds // 60}m {elapsed_seconds % 60}s"
                    else:
                        time_str = f"{elapsed_seconds // 3600}h {(elapsed_seconds % 3600) // 60}m"
                    
                    # Determine if the value is outdated (more than 2 seconds)
                    is_stale = elapsed_seconds > 2
                    
                    # Update the corresponding field
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
                
                # Update the values in the table
                self.tp2_tree.item(item_id, values=tuple(new_values))
                
                # Apply the corresponding tag to the row based on the update state
                if all_stale:
                    self.tp2_tree.item(item_id, tags=('stale',))
                else:
                    self.tp2_tree.item(item_id, tags=('active',))
    
    def reset_tp2_data(self):
        """Resets all TP2 data to its initial state"""
        for i in range(8):
            item_id = self.tp2_tree.get_children()[i]
            self.tp2_tree.item(item_id, values=(i, '--', 'Never', '--', 'Never', '--', 'Never', 'Never'), tags=('stale',))
            self.last_update_times[i] = {
                'R': None,
                'C': None,
                'O': None,
                'any': None
            }
            # Clear plotting data
            if i in self.plot_data:
                for angle_type in self.plot_data[i]:
                    self.plot_data[i][angle_type].clear()

    def open_plot_window(self):
        """Opens a new window with real-time plots for the selected group"""
        try:
            group_id = int(self.plot_group_combo.get())
            
            # If a window already exists for this group, bring it to front
            if group_id in self.plot_windows and self.plot_windows[group_id].window.winfo_exists():
                self.plot_windows[group_id].window.lift()
                return
            
            # Create a new plot window
            self.plot_windows[group_id] = PlotWindow(self.root, group_id, self)
            
        except Exception as e:
            messagebox.showerror("Plot Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = CanMonitorApp(root)
    # Add window close handler
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()