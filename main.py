

from language_map import LANGUAGE_MAP  # Import the mapping

import ctypes
import time
import win32process
import serial
import serial.tools.list_ports
import win32api
import win32gui

# State machine:
NONE = 1
INITIALIZE = 2
GET_LANG_STATE = 3
SEND_SERIAL_TO_ARDUINO = 4
GET_PORT_STATE_AND_ESTABLISH = 5
ERROR_STATE = 10

# Serial port definitions
BAUD_RATE = 9600
ARDUINO_PORT_DESCRIPTION = "USB-SERIAL CH340"
SERIAL_TIMEOUT = 0.01
KEEP_ALIVE_TIMER = 1

user32 = ctypes.WinDLL('user32', use_last_error=True)


# Get keyboard language
def get_current_keyboard_language():
    hwnd = win32gui.GetForegroundWindow()
    thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
    layout_id = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    lang_id = layout_id & 0xFFFF
    return LANGUAGE_MAP.get(lang_id, hex(lang_id))


def pc_increment_language_state():
    # Press Alt+Shift
    win32api.keybd_event(0x12, 0, 0, 0)  # Alt
    win32api.keybd_event(0x10, 0, 0, 0)  # Shift
    time.sleep(0.05)
    win32api.keybd_event(0x10, 0, 2, 0)  # Shift up
    win32api.keybd_event(0x12, 0, 2, 0)  # Alt up


def get_port_state_and_establish():
    status = "Unavailable"
    arduino_state = 0

    ports = serial.tools.list_ports.comports()
    if not ports:
        print("No serial ports found.")
        time.sleep(1)

    for port in ports:
        port_name = port.device
        description = port.description

        # Try opening the port to check if it's available
        try:
            if ARDUINO_PORT_DESCRIPTION in description:
                arduino_state = serial.Serial(port_name, BAUD_RATE, timeout=SERIAL_TIMEOUT)
                # Wait for serial connection stabilization
                time.sleep(2)
                if arduino_state:
                    status = "Available"
        except (serial.SerialException, OSError):
            status = "Busy or Unavailable"

        print(f"port state & establish {port_name} - {description} - {status} ")

    return status, arduino_state


def get_port_state():
    status = "Unavailable"
    ports = serial.tools.list_ports.comports()
    if not ports:
        print("No serial ports found.")
        time.sleep(1)

    for port in ports:
        port_name = port.device
        description = port.description

        if ARDUINO_PORT_DESCRIPTION in description:
            status = "Available"

    return status


def debug_print(debug_current_state_machine, debug_prev_state_machine, print_str):
    if debug_current_state_machine != debug_prev_state_machine:
        debug_prev_state_machine = debug_current_state_machine
        print(print_str)
    return debug_prev_state_machine


def monitor_language_and_send():
    prev_state_machine = NONE
    state_machine = INITIALIZE
    arduino_serial_conn = 0
    next_send = time.perf_counter()

    while state_machine != ERROR_STATE:
        if state_machine == INITIALIZE:
            # Debug prints
            prev_state_machine = debug_print(state_machine, prev_state_machine, "INITIALIZE")
            last_lang = None
            state_machine = GET_PORT_STATE_AND_ESTABLISH

        elif state_machine == GET_LANG_STATE:
            prev_state_machine = debug_print(state_machine, prev_state_machine, "GET_LANG_STATE")

            # Check serial port status
            state = get_port_state()
            if state == "Available":
                current_lang = get_current_keyboard_language()
                if current_lang != last_lang:
                    print(f"Language changed to: {current_lang}")
                    message = current_lang + "\n"
                    last_lang = current_lang
                    state_machine = SEND_SERIAL_TO_ARDUINO
                else:
                    # Language was not changed - Re-read it
                    state_machine = GET_LANG_STATE
            else:
                state_machine = GET_PORT_STATE_AND_ESTABLISH

        elif state_machine == SEND_SERIAL_TO_ARDUINO:
            # Debug prints
            prev_state_machine = debug_print(state_machine, prev_state_machine, "SEND_SERIAL_TO_ARDUINO")

            # Send language to Arduino port
            arduino_serial_conn.write(message.encode('utf-8'))
            print(f"NEW Language Sent to Arduino: {message.strip()}")
            state_machine = GET_LANG_STATE

        elif state_machine == GET_PORT_STATE_AND_ESTABLISH:
            # Debug prints
            prev_state_machine = debug_print(state_machine, prev_state_machine, "GET_PORT_STATE_AND_ESTABLISH")

            status, arduino_serial_conn = get_port_state_and_establish()
            if status == "Available":
                state_machine = GET_LANG_STATE
            else:
                state_machine = GET_PORT_STATE_AND_ESTABLISH

            print("Arduino connected status = ", status)

        else:
            print("State machine error")

        try:
            # Receive language change from Arduino
            state = get_port_state()
            if state == "Available" and arduino_serial_conn:
                line = arduino_serial_conn.readline().decode('utf-8').strip()
                if line == "LANGUAGE_TOGGLE":
                    print("Toggle = " + line)
                    pc_increment_language_state()

            # Send KEEP_ALIVE message to the Arduino side
            now = time.perf_counter()
            # Check if it's time to send the next message
            if now >= next_send:
                next_send = now + KEEP_ALIVE_TIMER
                if arduino_serial_conn:
                    arduino_serial_conn.write(b'KEEP_ALIVE\n')
        # Exception handling
        except Exception as e:
            print("Exception handling - ", e)
            last_lang = 0


monitor_language_and_send()
