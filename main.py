import time
import os

from multiprocessing import Process

import re

from PIL import ImageGrab, ImageOps
from easyocr import Reader
import numpy as np
import pyautogui
import keyboard as kb
import win32gui
import win32com.client as comctl
from datetime import datetime

keyboard = comctl.Dispatch("WScript.Shell")
reader = Reader(lang_list=['en'], gpu=True)

tools = {
    1: ['shovel', 'shouel', 'snovel', 'snouel', 'shavel', 'showvel', 'showel', 'shivel'],
    2: ['pickaxe', 'piokaxe', 'pickaxo', 'piokaxo', 'piskaxe'],
    3: ['axe', 'axo', 'sxe']
}
start_title = '1.19'

isProcessing = True

press_key = 4


def easyocr_read(obj):
    results = reader.readtext(
        image=obj,
        decoder='greedy',
        batch_size=10,
        low_text=0.3,
        text_threshold=0.5,
    )
    results = sorted(results, key=lambda x: x[0][0])
    text_results = [x[-2] for x in results]
    easy_output = " ".join(text_results)
    easy_output = easy_output.strip()
    easy_output = re.sub('\s{2,}', ' ', easy_output)

    return easy_output


def get_text(x, y, w, h):
    original_screenshot = ImageGrab.grab(bbox=(x + 120, y + 80, x + (w / 5 * 3) / 3 * 2, y + h / 3 * 2))
    # original_screenshot = ImageGrab.grab(bbox=(x + 100, y + 80, x + w - 100, y + h))

    screenshot = original_screenshot.copy()

    screenshot = ImageOps.grayscale(screenshot)
    screenshot = ImageOps.invert(screenshot)
    screenshot = screenshot.copy()

    screenshot_np = np.array(screenshot)

    # screenshot.show()

    start = datetime.now()

    text = easyocr_read(screenshot_np)
    is_done = False

    for key in tools.keys():
        value = tools[key]
        if is_done:
            break
        for name in value:
            if name in str(text).lower():
                is_done = True
                running_time = (datetime.now() - start).total_seconds()
                print(f'[macro] {name} (recognition: {value[0]}) '
                      f'running time: {running_time}s')
                if press_key != key:
                    keyboard.SendKeys(key, 0)
                    press_key = key

                if running_time > 1:
                    original_screenshot.save(f'.\\logs\\{datetime.now().timestamp()}.png', 'png')
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
        while isProcessing:
            get_text(x, y, w, h)


def change_tools():
    win32gui.EnumWindows(cb, None)


def hold_key():
    while isProcessing:
        pyautogui.keyDown('shift')


def detect_exit_key(processes):
    while True:
        if kb.is_pressed("ctrl"):
            isProcessing = False
            print('Exit...')
            for p in processes:
                p.terminate()
            os._exit(1)
            break


def main():
    p1 = Process(target=change_tools)
    p1.start()

    p2 = Process(target=hold_key)
    p2.start()

    detect_exit_key([p1, p2])


if __name__ == '__main__':
    main()
