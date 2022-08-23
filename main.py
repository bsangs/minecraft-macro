import time

from PIL import ImageGrab
import pytesseract
import win32gui
import win32com.client as comctl
from datetime import datetime

keyboard = comctl.Dispatch("WScript.Shell")
tools = {
    1: ['shovel', 'shouel', 'snovel', 'snouel', 'shavel'],
    2: ['pickaxe', 'piokaxe', 'pickaxo', 'piokaxo'],
    3: ['axe', 'axo', 'sxe']
}
start_title = '.png'
pytesseract.pytesseract.tesseract_cmd = r'tesseract'


def get_text(x, y, w, h):
    screenshot = ImageGrab.grab(bbox=(x + w / 4 * 3, y + h / 2, x + w, y + h - h/8))
    # screenshot.show()

    start = datetime.now()

    text: str = pytesseract.image_to_string(
        screenshot,
        lang='eng'
    )

    # print(text)
    is_done = False
    for key in tools.keys():
        value = tools[key]
        if is_done:
            break
        for name in value:
            if name in text:
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


while True:
    win32gui.EnumWindows(cb, None)
# time.sleep(1)

#
#
# screenshot = ImageGrab.grab(include_layered_windows=True)
#
# screenshot.show()

# filename = 'C:\\Users\\bsangs\\AppData\\Roaming\\.minecraft\\screenshots\\stone.png'
# image = Image.open(filename)

# start = datetime.now()
# text = pytesseract.image_to_string(
#     image,
#     lang='eng'
# )
#
# print(text, datetime.now() - start, 'ms')
