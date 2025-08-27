
from __future__ import annotations
import os
import ctypes
import time
import win32process
import serial
import serial.tools.list_ports
import win32api
import win32gui
from ctypes import wintypes

import json
from pathlib import Path
from typing import Dict, Optional, List


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
LANG_MAPPING_CHANGE_TIMER = 5

# GetLocaleInfoEx fields
LOCALE_ILANGUAGE            = 0x00000001  # hex LANGID string, e.g. "0409"
LOCALE_SENGLISHDISPLAYNAME  = 0x00000072  # "English (United States)"
LOCALE_SNAME                = 0x0000005C  # "en-US" (fallback)

BUF_LEN = 40  # LOCALE_NAME_MAX_LENGTH
# OUTPUT_PATH = Path("installed_languages.txt")  # Allocation mapping fIle name
OUTPUT_PATH = Path.home() / "Boten" / "installed_languages.txt"  # Allocation mapping fIle name
STATE_PATH = Path.home() / "Boten" / "color_allocations.json"

# WinAPI DLLs
user32   = ctypes.WinDLL("user32",   use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

# Signatures
user32.GetKeyboardLayoutList.argtypes = [wintypes.INT, ctypes.POINTER(ctypes.c_void_p)]
user32.GetKeyboardLayoutList.restype  = wintypes.UINT

kernel32.LCIDToLocaleName.argtypes = [wintypes.LCID, wintypes.LPWSTR, ctypes.c_int, wintypes.DWORD]
kernel32.LCIDToLocaleName.restype  = ctypes.c_int

kernel32.GetLocaleInfoEx.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.LPWSTR, ctypes.c_int]
kernel32.GetLocaleInfoEx.restype  = ctypes.c_int

# Configuration
COLOR_POOL: List[str] = ["Red", "Green", "Blue", "White", "Cyan", "Yellow", "Magenta"]

def _load_state() -> Dict[str, str]:
    # Ensure the directory exists; create if missing
    if not STATE_PATH.parent.exists():
        print(f"Creating directory: {STATE_PATH.parent}")
        STATE_PATH.parent.mkdir(parents=True, exist_ok=True)

    # Load the id→color mapping from disk; create file if missing or malformed.
    if not STATE_PATH.exists():
        _save_state({})
        return {}
    try:
        data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            # Keep only string->string pairs
            return {str(k): str(v) for k, v in data.items()}
    except Exception:
        pass
    # Malformed content: reset safely
    _save_state({})
    return {}

def _atomic_write(text: str) -> None:
    # Write atomically to reduce risk of partial writes
    tmp = STATE_PATH.with_suffix(".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(STATE_PATH)

def _save_state(mapping: Dict[str, str]) -> None:
    # Persist the id→color mapping
    _atomic_write(json.dumps(mapping, ensure_ascii=False, indent=2))

def allocate_color(identifier: str) -> Optional[str]:
    """
    Allocate a color for the given identifier.
    - If the identifier already has a color, return it.
    - Otherwise pick the first color from COLOR_POOL not currently assigned.
    - If none available, return None.
    """
    mapping = _load_state()
    if identifier in mapping:
        return mapping[identifier]

    used = set(mapping.values())
    for color in COLOR_POOL:
        if color not in used:
            mapping[identifier] = color
            _save_state(mapping)
            return color

    mapping[identifier] = None
    _save_state(mapping)
    return None  # Pool exhausted

def release_color(identifier: str) -> bool:
    """
    Release the color associated with the identifier.
    Returns True if released, False if the identifier was unknown.
    """
    mapping = _load_state()
    if identifier in mapping:
        mapping.pop(identifier)
        _save_state(mapping)
        return True
    return False

def _get_locale_info_ex(locale_name: str, field: int) -> str:
    buf = ctypes.create_unicode_buffer(BUF_LEN)
    n = kernel32.GetLocaleInfoEx(locale_name, field, buf, BUF_LEN)
    return buf.value if n > 0 else ""

def _lcid_to_locale_name(lcid: int) -> str:
    buf = ctypes.create_unicode_buffer(BUF_LEN)
    n = kernel32.LCIDToLocaleName(lcid, buf, BUF_LEN, 0)
    return buf.value if n > 0 else ""

def _installed_langids() -> list[int]:
    count = user32.GetKeyboardLayoutList(0, None)
    arr_type = ctypes.c_void_p * count
    arr = arr_type()
    user32.GetKeyboardLayoutList(count, arr)
    return sorted({int(hkl) & 0xFFFF for hkl in arr})

def language_color_allocation(lcid: int):
    # Retrieve the allocated color from file - keep Language color for-ever
    language_color = retrieve_saved_language_color(lcid)
    if language_color == "Language not found":
        allocated_new_color = allocate_color(lcid)
        language_color = allocated_new_color
        print("--- Allocate new color = ", allocated_new_color)

    return language_color

def build_lines() -> list[str]:
    lines: list[str] = []
    for lcid in _installed_langids():
        name = _lcid_to_locale_name(lcid)
        if not name:
            continue

        eng_display = _get_locale_info_ex(name, LOCALE_SENGLISHDISPLAYNAME)
        if not eng_display:
            eng_display = _get_locale_info_ex(name, LOCALE_SNAME)

        lang_color = language_color_allocation(lcid)
        lines.append(f"{lcid}:{eng_display}:{lang_color}")
    return lines

def save_language_color_mapping_if_changed() -> None:
    lines = build_lines()
    new_content = "\n".join(lines) + ("\n" if lines else "")
    old_content = ""
    if OUTPUT_PATH.exists():
        old_content = OUTPUT_PATH.read_text(encoding="utf-8")

    # Compute removed lines: lines that were in the old file but are not in the new lines
    old_lines = [ln for ln in old_content.splitlines() if ln.strip() != ""]
    new_set = set(ln for ln in lines if ln.strip() != "")
    removed_lines = [ln for ln in old_lines if ln not in new_set]
    if removed_lines:
        print("$$$ Removed line in file = ", removed_lines)

    # Extract the first token (ID) before the first ":" and call a handler with the list
    removed_ids = [ln.split(":", 1)[0] for ln in removed_lines if ":" in ln]
    if removed_ids:
        for ident in removed_ids:
            release_color(ident)

    if new_content != old_content:
        for line in lines:
            print(line)
        OUTPUT_PATH.write_text(new_content, encoding="utf-8")

def retrieve_saved_language_color(language_id: int):
    filename = OUTPUT_PATH
    default = "Language not found"

    if not OUTPUT_PATH.parent.exists():
        print(f"Creating directory: {OUTPUT_PATH.parent}")
        OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    # Ensure file exists
    if not os.path.exists(filename):
        open(filename, "a", encoding="utf-8").close()

    # search for key
    with open(filename, "r", encoding="utf-8") as f:
        for line in f:
            if line.startswith(f"{language_id}:"):
                return line.split(":", 2)[2].rstrip("\n")

    return default

# Get keyboard language
def get_current_keyboard_language():
    hwnd = win32gui.GetForegroundWindow()
    thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
    layout_id = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    lang_id = layout_id & 0xFFFF
    buf = ctypes.create_unicode_buffer(BUF_LEN)

    # Language string manipulation
    if kernel32.GetLocaleInfoW(lang_id, LOCALE_SENGLISHDISPLAYNAME, buf, BUF_LEN) > 0:
        string = buf.value

        # Retrieve/ Save language color
        language_color = retrieve_saved_language_color(lang_id)

        idx = string.find(" ")
        if idx != -1:
            result = language_color + ":" + string[:3] + string[idx:]
        else:
            result = string[:3]

        return result
    return ""

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
    lang_map_next_check = time.perf_counter()
    last_lang = None
    message = 0

    while state_machine != ERROR_STATE:
        if state_machine == INITIALIZE:
            # Debug prints
            prev_state_machine = debug_print(state_machine, prev_state_machine, "INITIALIZE")
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

        # Update Language to color mapping file
        # Check if it's time to check language mapping file should be updated
        now = time.perf_counter()
        if now >= lang_map_next_check:
            lang_map_next_check = now + LANG_MAPPING_CHANGE_TIMER
            save_language_color_mapping_if_changed()


monitor_language_and_send()
