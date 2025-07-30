

import ctypes
import time
import win32gui
import win32process
import serial
import serial.tools.list_ports

# Serial port definitions
SERIAL_PORT = 'COM3'     # Arduino port - TODO dynamic port name setting
BAUD_RATE = 9600

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
ERROR_STATE = 10


# Get keyboard language
def get_current_keyboard_language():
    hwnd = win32gui.GetForegroundWindow()
    thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
    layout_id = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    lang_id = layout_id & 0xFFFF
    return LANGUAGE_MAP.get(lang_id, hex(lang_id))


def monitor_language_and_send():

    state_machine = INITIALIZE

    while state_machine != ERROR_STATE:
        if state_machine == INITIALIZE:
            print("INITIALIZE")
            last_lang = None
            arduino = False
            state_machine = CHECK_SERIAL_CON_ESTABLISH

        elif state_machine == EST_SERIAL_CON:
            print("EST_SERIAL_CON")
            # Try to open serial connection
            try:
                arduino_serial = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=1)
                if not arduino_serial:
                    arduino_serial.open()
                    print("EST_SERIAL_CON - Try to open arduino = ", arduino_serial)

            # If exception occurred
            except serial.SerialException as e:
                print("❌ Serial error:", e)

            time.sleep(2)  # Wait 2 sec for establishing the connection
            if not arduino_serial:
                print("Fail to open Arduino serial connection", SERIAL_PORT)
                state_machine = EST_SERIAL_CON
            else:
                print("Serial connection established to Arduino on port", SERIAL_PORT)
                print("(Press Ctrl+C to stop)\n")
                state_machine = GET_LANG_STATE

        elif state_machine == GET_LANG_STATE:
            print("GET_LANG_STATE")
            current_lang = get_current_keyboard_language()
            if current_lang != last_lang:
                print(f"Language changed to: {current_lang}")
                message = current_lang + "\n"
                last_lang = current_lang
                state_machine = CHECK_SERIAL_CON_SEND_ARDUINO
            else:
                # Language was not changed - Re-read it
                state_machine = GET_LANG_STATE
                time.sleep(1)

        elif (state_machine == CHECK_SERIAL_CON_ESTABLISH or
                state_machine == CHECK_SERIAL_CON_SEND_ARDUINO or
                state_machine == CHECK_SERIAL_CON_ONLY):
            available_ports = [p.device for p in serial.tools.list_ports.comports()]
            if SERIAL_PORT in available_ports:
                print("CHECK_SERIAL_CON: port found on list:", SERIAL_PORT, "State = ", state_machine)
                # Handle next machine state based on the check type
                if state_machine == CHECK_SERIAL_CON_ESTABLISH:
                    state_machine = EST_SERIAL_CON  # Check + Establish
                elif state_machine == CHECK_SERIAL_CON_SEND_ARDUINO:
                    state_machine = SEND_SERIAL_TO_ARDUINO  # Check & send to Arduino - connection established
                else:
                    state_machine = CHECK_SERIAL_CON_ESTABLISH  # Check Only
                    time.sleep(1)
            else:
                print("CHECK_SERIAL_CON: port NOT found on list:", SERIAL_PORT)
                time.sleep(1)
                state_machine = CHECK_SERIAL_CON_ONLY

        elif state_machine == SEND_SERIAL_TO_ARDUINO:
            print("SEND_SERIAL_TO_ARDUINO")
            # Send language to Arduino port
            arduino_serial.write(message.encode('utf-8'))
            print(f"Sent to Arduino: {message.strip()}")
            state_machine = GET_LANG_STATE

        else:
            print("State machine error")


monitor_language_and_send()