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
    def __init__(self, parent, data_source):
        self.window = tk.Toplevel(parent)
        self.window.title("Real-time Angle Data Plot")
        self.window.geometry("1100x800")
        self.data_source = data_source
        self.time_window = 30  # seconds

        # Group/magnitude selection state
        self.group_vars = [tk.BooleanVar(value=True) for _ in range(8)]
        self.magnitude_vars = {
            'R': [tk.BooleanVar(value=True) for _ in range(8)],
            'C': [tk.BooleanVar(value=True) for _ in range(8)],
            'O': [tk.BooleanVar(value=True) for _ in range(8)],
        }

        self.setup_controls()
        self.setup_plots()

        self.ani = animation.FuncAnimation(
            self.fig, self.update_plots, interval=100, blit=False)

        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_controls(self):
        control_frame = ttk.Frame(self.window)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        # Group selection checkbuttons
        group_sel_frame = ttk.LabelFrame(control_frame, text="Groups")
        group_sel_frame.pack(side=tk.LEFT, padx=5, pady=5)
        for i in range(8):
            ttk.Checkbutton(
                group_sel_frame, text=f"G{i}", variable=self.group_vars[i],
                command=self.on_group_toggle
            ).grid(row=0, column=i, sticky=tk.W, padx=2)

        # Magnitude selection per group
        mag_sel_frame = ttk.LabelFrame(control_frame, text="Magnitudes")
        mag_sel_frame.pack(side=tk.LEFT, padx=5, pady=5)
        for i in range(8):
            col = i
            ttk.Label(mag_sel_frame, text=f"G{i}").grid(row=0, column=col)
            for j, mag in enumerate(['R', 'C', 'O']):
                ttk.Checkbutton(
                    mag_sel_frame, text=mag, variable=self.magnitude_vars[mag][i],
                    command=self.on_mag_toggle
                ).grid(row=j+1, column=col, sticky=tk.W)

        # Buttons to select/deselect all
        btns_frame = ttk.Frame(control_frame)
        btns_frame.pack(side=tk.LEFT, padx=10)
        ttk.Button(btns_frame, text="All Groups", command=self.select_all_groups).pack(fill=tk.X)
        ttk.Button(btns_frame, text="No Groups", command=self.deselect_all_groups).pack(fill=tk.X)
        ttk.Button(btns_frame, text="All Magnitudes", command=self.select_all_mags).pack(fill=tk.X)
        ttk.Button(btns_frame, text="No Magnitudes", command=self.deselect_all_mags).pack(fill=tk.X)

    def setup_plots(self):
        self.fig = Figure(figsize=(10, 7), dpi=100)
        self.axes = {}
        self.lines = {}
        mag_titles = {'R': 'Roll', 'C': 'Pitch', 'O': 'Orientation'}
        colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', '#ff8800']

        # One subplot per magnitude
        for idx, mag in enumerate(['R', 'C', 'O']):
            ax = self.fig.add_subplot(3, 1, idx+1)
            ax.set_title(mag_titles[mag])
            ax.set_ylabel("Degrees")
            ax.set_ylim(-180, 180)
            ax.grid(True)
            ax.axhline(y=0, color='k', linestyle='--', alpha=0.3)
            if mag == 'O':
                ax.set_xlabel("Time (seconds)")
            self.axes[mag] = ax
            self.lines[mag] = {}
            for group in range(8):
                # Each group gets a line per magnitude, now with dots
                line, = ax.plot([], [], color=colors[group % len(colors)], label=f"G{group}", marker='o')
                self.lines[mag][group] = line
            ax.legend(loc='upper right', fontsize='small', ncol=4)

        self.fig.tight_layout()
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.window)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def on_group_toggle(self):
        # If a group is disabled, also disable its magnitudes
        for i in range(8):
            if not self.group_vars[i].get():
                for mag in ['R', 'C', 'O']:
                    self.magnitude_vars[mag][i].set(False)
        self.canvas.draw_idle()

    def on_mag_toggle(self):
        # If any magnitude for a group is enabled, enable the group
        for i in range(8):
            if any(self.magnitude_vars[mag][i].get() for mag in ['R', 'C', 'O']):
                self.group_vars[i].set(True)
        self.canvas.draw_idle()

    def select_all_groups(self):
        for var in self.group_vars:
            var.set(True)
        for mag in ['R', 'C', 'O']:
            for var in self.magnitude_vars[mag]:
                var.set(True)
        self.canvas.draw_idle()

    def deselect_all_groups(self):
        for var in self.group_vars:
            var.set(False)
        for mag in ['R', 'C', 'O']:
            for var in self.magnitude_vars[mag]:
                var.set(False)
        self.canvas.draw_idle()

    def select_all_mags(self):
        for mag in ['R', 'C', 'O']:
            for var in self.magnitude_vars[mag]:
                var.set(True)
        for i in range(8):
            if any(self.magnitude_vars[mag][i].get() for mag in ['R', 'C', 'O']):
                self.group_vars[i].set(True)
        self.canvas.draw_idle()

    def deselect_all_mags(self):
        for mag in ['R', 'C', 'O']:
            for var in self.magnitude_vars[mag]:
                var.set(False)
        self.canvas.draw_idle()

    def update_plots(self, frame):
        now = time.time()
        for mag in ['R', 'C', 'O']:
            ax = self.axes[mag]
            for group in range(8):
                line = self.lines[mag][group]
                if self.group_vars[group].get() and self.magnitude_vars[mag][group].get():
                    data = self.data_source.get_plot_data(group)[mag]
                    filtered = [(t, v) for t, v in data if now - t <= self.time_window]
                    if filtered:
                        times, values = zip(*filtered)
                        rel_times = [now - t for t in times]
                        line.set_data(rel_times, values)
                        ax.set_xlim(self.time_window, 0)
                    else:
                        line.set_data([], [])
                else:
                    line.set_data([], [])
        return [self.lines[mag][g] for mag in ['R', 'C', 'O'] for g in range(8)]

    def on_close(self):
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
        
        # Continuous transmission variables
        self.continuous_active = False
        self.continuous_timer = None
        self.last_angle_data = {
            'group_id': 0,
            'angle_type': 'R',
            'angle_value': '0',
            'input_method': 'numeric',
            'angle_string': 'R0'
        }
        
        # Historical data for plotting
        self.plot_data = {}
        for i in range(8):
            self.plot_data[i] = {
                'R': deque(maxlen=500),  # Store up to 500 points
                'C': deque(maxlen=500),
                'O': deque(maxlen=500)
            }
        
        # Reference to plot window
        self.plot_window = None
        
        # Variables for random transmission
        self.random_transmission_active = False
        self.random_transmission_thread = None
        
        # Variables for search functionality
        self.search_term = ""
        self.search_matches = []
        self.current_match = -1
        
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
        
        # Send button (shared) - modify this section
        send_controls_frame = ttk.Frame(presets_frame)
        send_controls_frame.grid(row=3, column=0, columnspan=4, sticky=tk.W, padx=5, pady=5)
        
        self.send_preset_btn = ttk.Button(send_controls_frame, text="Send Angle", command=self.send_tp2_angle)
        self.send_preset_btn.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        # Add continuous transmission toggle
        self.continuous_var = tk.BooleanVar(value=False)
        self.continuous_check = ttk.Checkbutton(
            send_controls_frame, 
            text="Send Continuously", 
            variable=self.continuous_var,
            command=self.toggle_continuous_transmission
        )
        self.continuous_check.grid(row=0, column=1, padx=5, pady=5)
        
        # Add period selection
        ttk.Label(send_controls_frame, text="Period (ms):").grid(row=0, column=2, padx=5, pady=5)
        
        # Available periods from 50ms to 2s
        periods = [50, 100, 200, 500, 1000, 2000]
        self.period_combo = ttk.Combobox(
            send_controls_frame, 
            width=5, 
            values=[str(p) for p in periods],
            state="readonly"
        )
        self.period_combo.grid(row=0, column=3, padx=5, pady=5)
        self.period_combo.current(2)  # Default to 200ms
        
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

        # Multi-group selection for random transmission
        ttk.Label(random_frame, text="Source Groups:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.random_group_vars = [tk.BooleanVar(value=False) for _ in range(8)]
        group_sel_frame = ttk.Frame(random_frame)
        group_sel_frame.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        for i in range(8):
            ttk.Checkbutton(
                group_sel_frame, text=f"G{i}", variable=self.random_group_vars[i]
            ).grid(row=0, column=i, sticky=tk.W, padx=2)

        # Buttons to select/deselect all groups
        group_btns_frame = ttk.Frame(random_frame)
        group_btns_frame.grid(row=0, column=2, sticky=tk.W, padx=5)
        ttk.Button(group_btns_frame, text="All", command=lambda: [v.set(True) for v in self.random_group_vars]).pack(fill=tk.X)
        ttk.Button(group_btns_frame, text="None", command=lambda: [v.set(False) for v in self.random_group_vars]).pack(fill=tk.X)

        # Create a compact mode selection with a legend
        ttk.Label(random_frame, text="Signal Mode:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=(8,2))
        mode_opts = ["Sine", "Const", "Noise"]
        self.random_group_mode = []
        
        # Create a frame with 2 rows of mode selectors (4 per row)
        mode_sel_frame = ttk.Frame(random_frame)
        mode_sel_frame.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5, pady=(8,2))
        
        for i in range(8):
            var = tk.StringVar(value="Sine")
            row = i // 4  # Put 4 groups on each row
            col = i % 4   # Arrange in columns within each row
            
            group_frame = ttk.LabelFrame(mode_sel_frame, text=f"G{i}", padding=(2,0))
            group_frame.grid(row=row, column=col, padx=3, pady=2)
            
            cb = ttk.Combobox(group_frame, width=5, values=mode_opts, 
                              state="readonly", textvariable=var)
            cb.pack(padx=1, pady=1)
            self.random_group_mode.append(var)

        # Add a legend for the modes
        legend_frame = ttk.LabelFrame(random_frame, text="Mode Legend", padding=(5,0))
        legend_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=5, pady=(5,2))
        ttk.Label(legend_frame, text="Sine: Smooth sinusoidal signal").pack(anchor=tk.W)
        ttk.Label(legend_frame, text="Const: Fixed value signal").pack(anchor=tk.W) 
        ttk.Label(legend_frame, text="Noise: Random fluctuating signal").pack(anchor=tk.W)

        # Status label for timing info
        self.random_status = ttk.Label(random_frame, text="Idle", width=60)
        self.random_status.grid(row=3, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)

        # Start/Stop button
        self.random_btn = ttk.Button(random_frame, text="Start Random Transmission",
                                    command=self.toggle_random_transmission)
        self.random_btn.grid(row=4, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # Add a separator
        ttk.Separator(random_frame, orient=tk.HORIZONTAL).grid(
            row=5, column=0, columnspan=3, sticky=tk.EW, pady=10)
        
        # # Information about timing rules
        # timing_text = ("Timing rules:\n"
        #               "• Max 20 packets/second per group\n"
        #               "• Send immediately if angle changes ≥5°\n"
        #               "• Send at least every 2 seconds\n"
        #               "• Roll, pitch and orientation treated independently")
        # ttk.Label(random_frame, text=timing_text, justify=tk.LEFT).grid(
        #     row=6, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
        
        # === RIGHT COLUMN ===
        right_frame = ttk.LabelFrame(main_frame, text="CAN Messages", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Search frame for message filtering
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=5)
        self.search_entry = ttk.Entry(search_frame, width=20)
        self.search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.search_entry.bind("<Return>", self.search_text)
        
        self.search_btn = ttk.Button(search_frame, text="Find", command=self.search_text)
        self.search_btn.pack(side=tk.LEFT, padx=5)
        
        self.prev_btn = ttk.Button(search_frame, text="↑", width=2, command=lambda: self.navigate_search(-1))
        self.prev_btn.pack(side=tk.LEFT, padx=2)
        
        self.next_btn = ttk.Button(search_frame, text="↓", width=2, command=lambda: self.navigate_search(1))
        self.next_btn.pack(side=tk.LEFT, padx=2)
        
        self.match_label = ttk.Label(search_frame, text="")
        self.match_label.pack(side=tk.LEFT, padx=5)
        
        # Area to display received messages
        self.rx_text = scrolledtext.ScrolledText(right_frame, width=50, height=20)
        self.rx_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Configure colors for messages
        self.rx_text.tag_config("tx_msg", foreground="green")
        self.rx_text.tag_config("rx_msg", foreground="blue")
        self.rx_text.tag_config("system", foreground="black")
        self.rx_text.tag_config("error", foreground="red")
        self.rx_text.tag_config("timestamp", foreground="gray")
        self.rx_text.tag_config("search_highlight", background="yellow")
        
        # Create right-click (context) menu for copy functionality
        self.context_menu = tk.Menu(self.rx_text, tearoff=0)
        self.context_menu.add_command(label="Copy Selected", command=self.copy_selected)
        self.context_menu.add_command(label="Copy All", command=self.copy_all)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Clear All", command=lambda: self.rx_text.delete(1.0, tk.END))
        
        # Bind right-click to show context menu
        self.rx_text.bind("<Button-3>", self.show_context_menu)
        
        # Buttons to clear and enable/disable autoscroll
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.clear_btn = ttk.Button(btn_frame, text="Clear", command=lambda: self.rx_text.delete(1.0, tk.END))
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Add autoscroll toggle
        self.autoscroll_var = tk.BooleanVar(value=True)
        self.autoscroll_check = ttk.Checkbutton(
            btn_frame, text="Autoscroll", variable=self.autoscroll_var)
        self.autoscroll_check.pack(side=tk.LEFT, padx=5)
        
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
        self.open_plot_btn = ttk.Button(plotting_frame, text="Open Plot Window", 
                                        command=self.open_plot_window)
        self.open_plot_btn.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

    def format_timestamp(self):
        """Returns a formatted timestamp string for the current time"""
        now = datetime.now()
        return now.strftime("[%H:%M:%S.%f")[:-3] + "]"  # Format as [HH:MM:SS.mmm]

    def toggle_input_method(self):
        """Toggles between numeric and string input methods"""
        if self.input_method.get() == "numeric":
            self.string_frame.grid_remove()
            self.numeric_frame.grid()
        else:
            self.numeric_frame.grid_remove()
            self.string_frame.grid()
    
    def autoscroll(self):
        """Only scrolls to the end if autoscroll is enabled"""
        if self.autoscroll_var.get():
            self.rx_text.see(tk.END)
    
    def toggle_continuous_transmission(self):
        """Starts or stops continuous angle transmission"""
        if self.continuous_var.get():
            # Check if we're connected
            if not self.is_connected:
                messagebox.showwarning("Not Connected", "Connect to the serial port first")
                self.continuous_var.set(False)
                return
            
            # Start continuous transmission
            self.continuous_active = True
            self.period_combo.configure(state="disabled")  # Disable changing period while active
            self.send_continuous_angle()
            timestamp = self.format_timestamp()
            self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
            self.rx_text.insert(tk.END, f"Started continuous angle transmission ({self.period_combo.get()}ms)\n", "system")
            self.autoscroll()
        else:
            # Stop continuous transmission
            self.continuous_active = False
            self.period_combo.configure(state="readonly")  # Re-enable period selection
            if self.continuous_timer:
                self.root.after_cancel(self.continuous_timer)
                self.continuous_timer = None
            timestamp = self.format_timestamp()
            self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
            self.rx_text.insert(tk.END, "Stopped continuous angle transmission\n", "system")
            self.autoscroll()

    def send_continuous_angle(self):
        """Sends the last angle continuously at the selected period"""
        if not self.continuous_active or not self.is_connected:
            return
        
        try:
            # Send the angle using the last stored values
            group_id = self.last_angle_data['group_id']
            can_id = f"{0x100 + group_id:x}"
            
            # Handle based on which input method was last used
            if self.last_angle_data['input_method'] == "numeric":
                angle_type = self.last_angle_data['angle_type']
                angle_value = self.last_angle_data['angle_value']
                
                # Convert angle type and value to hexadecimal bytes
                data_bytes = []
                data_bytes.append(f"{ord(angle_type):02x}")
                for char in angle_value:
                    data_bytes.append(f"{ord(char):02x}")
                
                display_msg = f"Continuous: {angle_type}={angle_value}° (Group {group_id})"
                
            else:  # string format
                angle_string = self.last_angle_data['angle_string']
                
                # Convert the entire string to hex bytes
                data_bytes = []
                for char in angle_string:
                    data_bytes.append(f"{ord(char):02x}")
                
                display_msg = f"Continuous: {angle_string} (Group {group_id})"
            
            # Build CAN command
            cmd = f"SEND_{can_id}"
            for byte in data_bytes:
                cmd += f"_{byte}"
            
            # Send the command
            self.serial_port.write((cmd + "\n").encode('utf-8'))
            
            # Periodically log the continuous transmission (once every ~2 seconds)
            current_time = time.time()
            if not hasattr(self, 'last_continuous_log') or current_time - self.last_continuous_log >= 2.0:
                timestamp = self.format_timestamp()
                self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
                self.rx_text.insert(tk.END, f"{display_msg}\n", "tx_msg")
                self.autoscroll()
                self.last_continuous_log = current_time
            
            # Schedule the next transmission
            period = int(self.period_combo.get())
            self.continuous_timer = self.root.after(period, self.send_continuous_angle)
            
        except Exception as e:
            self.rx_text.insert(tk.END, f"Error in continuous transmission: {str(e)}\n", "error")
            self.autoscroll()
            self.continuous_var.set(False)
            self.toggle_continuous_transmission()  # Stop continuous transmission
    
    def toggle_random_transmission(self):
        """Starts or stops the random transmission mode for selected groups"""
        if not self.is_connected:
            messagebox.showwarning("Not Connected", "Connect to the serial port first")
            return

        selected_groups = [i for i, v in enumerate(self.random_group_vars) if v.get()]
        if not self.random_transmission_active:
            if not selected_groups:
                messagebox.showwarning("No Groups Selected", "Select at least one group to start random transmission.")
                return
            self.random_transmission_active = True
            self.random_btn.config(text="Stop Random Transmission")
            self.random_status.config(text="Starting...")

            # Reset per-group state
            self.random_group_state = {}
            for group_id in selected_groups:
                mode = self.random_group_mode[group_id].get()
                self.random_group_state[group_id] = {
                    'last_values': {'R': 0, 'C': 0, 'O': 0},
                    'last_sent_time': {'R': 0, 'C': 0, 'O': 0},
                    'mode': mode,
                    'const_value': random.randint(-90, 90),  # Para modo constante
                    'sine_params': {
                        'R': {
                            'amplitude': random.randint(50, 120),
                            'period': random.uniform(0.1, 10),
                            'phase': random.uniform(0, 2*math.pi),
                            'offset': random.randint(-50, 50)
                        },
                        'C': {
                            'amplitude': random.randint(50, 120),
                            'period': random.uniform(0.1, 10),
                            'phase': random.uniform(0, 2*math.pi),
                            'offset': random.randint(-50, 50)
                        },
                        'O': {
                            'amplitude': random.randint(50, 120),
                            'period': random.uniform(0.1, 10),
                            'phase': random.uniform(0, 2*math.pi),
                            'offset': random.randint(-50, 50)
                        }
                    },
                    'start_time': time.time()
                }
            self.random_transmission_thread = threading.Thread(
                target=self.random_transmission_loop_multi,
                args=(selected_groups,),
                daemon=True
            )
            self.random_transmission_thread.start()
            timestamp = self.format_timestamp()
            self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
            self.rx_text.insert(tk.END, f"Started random transmission for Groups: {', '.join(str(g) for g in selected_groups)}\n", "system")
            self.autoscroll()
        else:
            self.random_transmission_active = False
            self.random_btn.config(text="Start Random Transmission")
            self.random_status.config(text="Idle")
            timestamp = self.format_timestamp()
            self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
            self.rx_text.insert(tk.END, "Stopped random transmission\n", "system")
            self.autoscroll()

    def random_transmission_loop_multi(self, group_ids):
        """Thread function to send random angle values for multiple groups with TP2 timing rules"""
        angle_types = ['R', 'C', 'O']
        next_angle_index = {g: 0 for g in group_ids}
        while self.random_transmission_active and self.is_connected:
            now = time.time()
            for group_id in group_ids:
                state = self.random_group_state[group_id]
                elapsed = now - state['start_time']
                angle_type = angle_types[next_angle_index[group_id]]
                next_angle_index[group_id] = (next_angle_index[group_id] + 1) % 3

                mode = state.get('mode', 'Sine')
                if mode == "Sine":
                    params = state['sine_params'][angle_type]
                    sine_value = params['amplitude'] * math.sin(2 * math.pi * elapsed / params['period'] + params['phase']) + params['offset']
                    new_value = int(sine_value + random.uniform(-2, 2))
                elif mode == "Const":
                    new_value = state['const_value']
                elif mode == "Noise":
                    new_value = random.randint(-179, 180)
                else:
                    new_value = 0
                new_value = max(-179, min(180, new_value))

                # Timing logic
                last_sent = state['last_sent_time'].get(angle_type, 0)
                last_val = state['last_values'].get(angle_type, 0)
                should_send = False
                reason = ""
                if abs(new_value - last_val) >= 5:
                    should_send = True
                    reason = "≥5° change"
                elif now - last_sent >= 2.0:
                    should_send = True
                    reason = "2s timeout"
                # Max 5 packets/sec per group per angle type (0.500s)
                if should_send and (now - last_sent) >= 0.500:
                    can_id = f"{0x100 + group_id:x}"
                    angle_string = f"{angle_type}{new_value}"
                    data_bytes = [f"{ord(c):02x}" for c in angle_string]
                    cmd = f"SEND_{can_id}" + ''.join(f"_{b}" for b in data_bytes)
                    try:
                        self.serial_port.write((cmd + "\n").encode('utf-8'))
                        status_text = f"G{group_id} {angle_type}={new_value}° ({reason}, {mode})"
                        self.root.after(0, lambda t=status_text: self.random_status.config(text=t))
                        state['last_sent_time'][angle_type] = now
                        state['last_values'][angle_type] = new_value
                        msg = f"Random: Sent {angle_type}={new_value}° for Group {group_id} ({reason}, {mode})"
                        timestamp = self.format_timestamp()
                        self.root.after(0, lambda t=timestamp, m=msg: self.rx_text.insert(tk.END, f"{t} ", "timestamp"))
                        self.root.after(0, lambda m=msg: self.rx_text.insert(tk.END, f"{m}\n", "tx_msg"))
                        self.root.after(0, self.autoscroll)
                        time.sleep(0.01)
                    except Exception as e:
                        self.root.after(0, lambda: self.rx_text.insert(
                            tk.END, f"Error sending angle: {str(e)}\n", "error"))
                        self.root.after(0, self.autoscroll)
                        if not self.is_connected:
                            self.random_transmission_active = False
                            break
                # Small sleep to prevent CPU overuse
                time.sleep(0.005)
            # If no groups are selected anymore, stop
            if not any(self.random_group_vars[g].get() for g in group_ids):
                self.root.after(0, self.toggle_random_transmission)
                break

    def on_closing(self):
        """Cleanup when the application is closing"""
        # Stop continuous transmission if active
        if self.continuous_active:
            self.continuous_active = False
            if self.continuous_timer:
                self.root.after_cancel(self.continuous_timer)
        
        # Stop random transmission if active
        self.random_transmission_active = False
        
        # Disconnect if connected
        if self.is_connected:
            self.toggle_connection()
        
        # Close plot window
        if self.plot_window and hasattr(self.plot_window, 'window') and self.plot_window.window.winfo_exists():
            self.plot_window.window.destroy()
        
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
                self.serial_port = serial.Serial(port, 921600, timeout=1)
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
                
                # Display system information with timestamps
                os_info = platform.platform()
                timestamp = self.format_timestamp()
                self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
                self.rx_text.insert(tk.END, f"System: {os_info}\n", "system")
                
                timestamp = self.format_timestamp()
                self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
                self.rx_text.insert(tk.END, f"Connected to {port} @ 115200 bps\n", "system")
                self.autoscroll()
            except Exception as e:
                messagebox.showerror("Connection Error", str(e))
        else:
            # If continuous transmission is active, stop it
            if self.continuous_active:
                self.continuous_var.set(False)
                self.continuous_active = False
                if self.continuous_timer:
                    self.root.after_cancel(self.continuous_timer)
                    self.continuous_timer = None
            
            self.should_read = False
            if self.serial_port:
                self.serial_port.close()
            self.is_connected = False
            self.connect_btn['text'] = "Connect"
            timestamp = self.format_timestamp()
            self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
            self.rx_text.insert(tk.END, "Disconnected\n", "system")
            self.autoscroll()
            
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
                    self.root.after(0, self.autoscroll)
            time.sleep(0.01)
    
    def process_received_data(self, data):
        """Processes data received via serial"""
        if not data:
            return
        
        # Add timestamp to the message
        timestamp = self.format_timestamp()
        
        # Insert timestamp with gray color, then the message with blue color
        self.root.after(0, lambda: self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp"))
        self.root.after(0, lambda: self.rx_text.insert(tk.END, f"{data}\n", "rx_msg"))
        
        # Only auto-scroll if the autoscroll option is enabled
        self.root.after(0, self.autoscroll)
        
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
            # Green color for sent messages with timestamp
            timestamp = self.format_timestamp()
            self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
            self.rx_text.insert(tk.END, f"Sending: {cmd}\n", "tx_msg")
            self.autoscroll()
            
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
            
            # Store the group ID for continuous transmission
            self.last_angle_data['group_id'] = group_id
            
            # Handle based on selected input method
            if self.input_method.get() == "numeric":
                # Original numerical input handling
                angle_type = self.angle_type.get()
                angle_value = self.angle_value.get()
                
                # Store values for continuous transmission
                self.last_angle_data['input_method'] = 'numeric'
                self.last_angle_data['angle_type'] = angle_type
                self.last_angle_data['angle_value'] = angle_value
                
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
                
                # Store values for continuous transmission
                self.last_angle_data['input_method'] = 'string'
                self.last_angle_data['angle_string'] = angle_string
                
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
            # Green color for sent messages with timestamp
            timestamp = self.format_timestamp()
            self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
            self.rx_text.insert(tk.END, f"{display_msg}\n", "tx_msg")
            self.autoscroll()
            
            # If continuous transmission is active, restart it with the new values
            if self.continuous_active:
                if self.continuous_timer:
                    self.root.after_cancel(self.continuous_timer)
                self.last_continuous_log = 0  # Force log the first message
                self.continuous_timer = self.root.after(int(self.period_combo.get()), self.send_continuous_angle)
                
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
            # Green color for sent messages with timestamp
            timestamp = self.format_timestamp()
            self.rx_text.insert(tk.END, f"{timestamp} ", "timestamp")
            self.rx_text.insert(tk.END, f"Changing CAN mode: {mode}\n", "tx_msg")
            self.autoscroll()
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
        """Opens the single real-time plot window (all groups/magnitudes)"""
        try:
            if self.plot_window and self.plot_window.window.winfo_exists():
                self.plot_window.window.lift()
                return
            self.plot_window = PlotWindow(self.root, self)
        except Exception as e:
            messagebox.showerror("Plot Error", str(e))

    def show_context_menu(self, event):
        """Show the context menu on right-click"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            # Make sure to release the menu
            self.context_menu.grab_release()
    
    def copy_selected(self):
        """Copy selected text to clipboard"""
        try:
            selected_text = self.rx_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
        except tk.TclError:
            # No selection
            pass
    
    def copy_all(self):
        """Copy all text to clipboard"""
        all_text = self.rx_text.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(all_text)
        
    def search_text(self, event=None):
        """Search for the given text in the message display"""
        search_term = self.search_entry.get().strip()
        if not search_term:
            self.match_label.config(text="")
            self.clear_search_highlights()
            return
        
        self.search_term = search_term
        self.search_matches = []
        self.current_match = -1
        
        # Clear previous highlights
        self.clear_search_highlights()
        
        # Get all text
        all_text = self.rx_text.get(1.0, tk.END)
        
        # Find all occurrences of the search term
        start_idx = 0
        while True:
            idx = all_text.find(self.search_term, start_idx)
            if idx == -1:
                break
                
            line_count = all_text[:idx].count('\n')
            line_start = all_text[:idx].rfind('\n')
            if line_start == -1:
                line_start = 0
            else:
                line_start += 1  # Skip the newline character
                
            col = idx - line_start
            
            # tkinter text indices are 1-based for lines
            self.search_matches.append(f"{line_count + 1}.{col}")
            start_idx = idx + len(self.search_term)
        
        # Update match count label
        if self.search_matches:
            self.match_label.config(text=f"1/{len(self.search_matches)}")
            self.current_match = 0
            self.highlight_current_match()
        else:
            self.match_label.config(text="No matches")
    
    def clear_search_highlights(self):
        """Clear all search highlights"""
        self.rx_text.tag_remove("search_highlight", "1.0", tk.END)
    
    def highlight_current_match(self):
        """Highlight the current match"""
        if not self.search_matches or self.current_match < 0:
            return
            
        # Get the position of the current match
        pos = self.search_matches[self.current_match]
        end_idx = f"{pos}+{len(self.search_term)}c"
        
        # Highlight the match
        self.rx_text.tag_add("search_highlight", pos, end_idx)
        
        # Ensure the match is visible
        self.rx_text.see(pos)
        
        # Update the count label
        self.match_label.config(text=f"{self.current_match + 1}/{len(self.search_matches)}")
    
    def navigate_search(self, direction):
        """Navigate through search results (1 = forward, -1 = backward)"""
        if not self.search_matches:
            return
            
        self.current_match = (self.current_match + direction) % len(self.search_matches)
        
        # Clear previous highlights and highlight the current match
        self.clear_search_highlights()
        self.highlight_current_match()

if __name__ == "__main__":
    root = tk.Tk()

    try:
        # Get the directory of the current script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "icon.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        print(f"Could not set icon: {e}")

    app = CanMonitorApp(root)
    # Add window close handler
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()