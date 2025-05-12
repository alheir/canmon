# canmon - CAN Monitor and Validator for TP2

`canmon` is a tool designed to help students of the ITBA Embedded Systems course validate and debug their implementation of the CAN protocol for Practical Assignment No. 2. It consists of an Arduino firmware and a graphical user interface (GUI) in Python.

## Main Components

1.  **Arduino Firmware (`fw/arducanmon/`)**:
    *   Resides on an Arduino Uno equipped with an MCP2515 CAN module.
    *   Acts as a gateway, translating commands from the PC (via serial) to CAN messages, and vice versa.

2.  **Graphical User Interface (`gui/`)**:
    *   Desktop application developed in Python with Tkinter and Matplotlib.
    *   Allows visualizing CAN messages in real time, sending custom and TP2-specific messages.
    *   Includes a table to monitor the angles reported by each group and real-time graphs.

3.  **TP2 Assignment (`tp2/`)**:
    *   Contains the file `tp2/assignment2025.txt/.pdf` with the description and requirements of Practical Assignment No. 2.

## Repository Structure

```
├── fw/
│   └── arducanmon/  # Arduino firmware source code
│       ├── platformio.ini
│       └── src/
│           └── main.cpp
├── gui/  # Python GUI source code
│   ├── main.py
│   └── requirements.txt
├── tp2/
│   └── assignment.txt  # TP2 assignment description
├── LICENSE  # Project license
└── README.md  # This file
```
