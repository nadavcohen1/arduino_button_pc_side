

import ctypes
import time
import win32gui
import win32process
import serial
import serial.tools.list_ports
import win32api
import win32gui


# Serial port definitions
SERIAL_PORT = 'COM4'    # Arduino port - TODO dynamic port name setting
BAUD_RATE = 9600
ARDUINO_PORT_DESCRIPTION = "USB-SERIAL CH340 (COM4)"

user32 = ctypes.WinDLL('user32', use_last_error=True)

# Language mapping identification - TODO Add all Languages
LANGUAGE_MAP = {
    0x0409: 'EN',  # English
    0x040D: 'HE',  # Hebrew
    0x040C: 'FR',  # French
    0x0410: 'IT',  # Italian
    0x0419: 'RU',  # Russian
    0x0411: 'JA',  # Japanese
}

# State machine:
INITIALIZE = 1
CHECK_SERIAL_CON_ESTABLISH = 2      # Check and Establish
CHECK_SERIAL_CON_SEND_ARDUINO = 3   # Check and send to Arduino
CHECK_SERIAL_CON_ONLY = 4           # Check only
EST_SERIAL_CON = 5
GET_LANG_STATE = 6
SEND_SERIAL_TO_ARDUINO = 7
GET_PORT_STATE_AND_ESTABLISH = 8
ERROR_STATE = 10


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

    for port in ports:
        port_name = port.device
        description = port.description

        # Try opening the port to check if it's available
        try:
            arduino_state = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=1)
            if arduino_state and description == ARDUINO_PORT_DESCRIPTION:
                status = "Available"
            # with serial.Serial(port_name, baudrate=9600, timeout=1) as ser:
            #    if description == ARDUINO_PORT_DESCRIPTION:
            #        status = "Available"
        except (serial.SerialException, OSError):
            status = "Busy or Unavailable"

        print(f"port state & establish {port_name} - {description} ")

    return status, arduino_state


def get_port_state():
    status = "Unavailable"
    ports = serial.tools.list_ports.comports()
    if not ports:
        print("No serial ports found.")

    for port in ports:
        port_name = port.device
        description = port.description

        if description == "USB-SERIAL CH340 (COM4)":
            status = "Available"

        print(f"get port state = {port_name} - {description} - {status}")

    return status


def monitor_language_and_send():

    state_machine = INITIALIZE
    arduino_serial_conn = 0

    while state_machine != ERROR_STATE:
        if state_machine == INITIALIZE:
            print("INITIALIZE")
            last_lang = None
            state_machine = GET_PORT_STATE_AND_ESTABLISH

        elif state_machine == GET_LANG_STATE:
            print("GET_LANG_STATE")
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
                    time.sleep(1)
            else:
                state_machine = GET_PORT_STATE_AND_ESTABLISH

        elif state_machine == SEND_SERIAL_TO_ARDUINO:
            print("SEND_SERIAL_TO_ARDUINO")
            # Check serial port status
            state = get_port_state()
            if state == "Available":
                # Send language to Arduino port
                arduino_serial_conn.write(message.encode('utf-8'))
                print(f"Sent to Arduino: {message.strip()}")
                state_machine = GET_LANG_STATE
            else:
                print("NO Arduino port available")
                state_machine = GET_PORT_STATE_AND_ESTABLISH

        elif state_machine == GET_PORT_STATE_AND_ESTABLISH:
            print("GET_PORT_STATE_AND_ESTABLISH")
            status, arduino_serial_conn = get_port_state_and_establish()
            if status == "Available":
                state_machine = GET_LANG_STATE
            else:
                state_machine = GET_PORT_STATE_AND_ESTABLISH

            print("Arduino connected status = ", status)
            time.sleep(1)

        else:
            print("State machine error")

        # Receive language change from Arduino
        state = get_port_state()
        if state == "Available" and arduino_serial_conn:
            line = arduino_serial_conn.readline().decode('utf-8').strip()
            if line == "LANGUAGE_TOGGLE":
                print("Toggle = " + line)
                pc_increment_language_state()


monitor_language_and_send()
