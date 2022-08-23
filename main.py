import time
import os

import threading

from PIL import ImageGrab, ImageOps
import pytesseract
import pyautogui
import keyboard as kb
import win32gui
import win32com.client as comctl
from datetime import datetime

keyboard = comctl.Dispatch("WScript.Shell")
tools = {
    1: ['shovel', 'shouel', 'snovel', 'snouel', 'shavel', 'showvel', 'showel', 'shivel'],
    2: ['pickaxe', 'piokaxe', 'pickaxo', 'piokaxo', 'piskaxe'],
    3: ['axe', 'axo', 'sxe']
}
start_title = '1.19'
pytesseract.pytesseract.tesseract_cmd = r'tesseract'

isProcessing = True


def get_text(x, y, w, h):
    screenshot = ImageGrab.grab(bbox=(x + 120, y + 80, x + (w / 5 * 3) / 3 * 2, y + h / 3 * 2))

    screenshot = ImageOps.grayscale(screenshot)
    screenshot = ImageOps.invert(screenshot)
    screenshot = screenshot.copy()

    # screenshot.show()

    start = datetime.now()

    text = pytesseract.image_to_string(
        screenshot,
        lang='eng'
    )
    # print("text", text)
    is_done = False
    for key in tools.keys():
        value = tools[key]
        if is_done:
            break
        for name in value:
            if name in str(text).lower():
                is_done = True
                print(f'[macro] {name} (recognition: {value[0]}) '
                      f'running time: {(datetime.now() - start).total_seconds()}s')
                keyboard.SendKeys(key, 0)
                break


def cb(hwnd, extra):
    rect = win32gui.GetWindowRect(hwnd)
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y

    win_title: str = win32gui.GetWindowText(hwnd)

    is_visible = win32gui.IsWindowVisible(hwnd)
    if w + h == 0 or not is_visible or len(win_title) == 0:
        return

    if start_title in win_title:
        get_text(x, y, w, h)


def change_tools():
    while isProcessing:
        win32gui.EnumWindows(cb, None)


def hold_key():
    while isProcessing:
        pyautogui.keyDown('shift')


def detect_exit_key():
    while True:
        if kb.is_pressed("ctrl"):
            isProcessing = False
            print('Exit...')
            os._exit(1)
            break


detect_exit_key_thread = threading.Thread(target=detect_exit_key)
detect_exit_key_thread.start()

for i in range(5):
    print("[macro] Starting %d sec left." % (5 - i))
    time.sleep(1)

hold_key_thread = threading.Thread(target=hold_key)
hold_key_thread.start()

change_tools_thread = threading.Thread(target=change_tools)
change_tools_thread.start()
