import datetime
import inspect
import os
import random
import re
import sys
import string
import time
from mss.base import MSSBase
import win32gui
import win32con
from mss import mss
from typing import Optional

PASS = "PASS"
FAIL = "FAIL"


text_elements = (
    "GuiTextField", 
    "GuiCTextField", 
    "GuiPasswordField", 
    "GuiLabel", 
    "GuiTitlebar", 
    "GuiStatusbar", 
    "GuiButton", 
    "GuiTab", 
    "GuiShell", 
    "GuiStatusPane")

vkeys = ["ENTER", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
    None, "SHIFT+F2", "SHIFT+F3", "SHIFT+F4", "SHIFT+F5", "SHIFT+F6", "SHIFT+F7", "SHIFT+F8",
    "SHIFT+F9", "CTRL+SHIFT+0", "SHIFT+F11", "SHIFT+F12", "CTRL+F1", "CTRL+F2", "CTRL+F3", "CTRL+F4",
    "CTRL+F5", "CTRL+F6", "CTRL+F7", "CTRL+F8", "CTRL+F9", "CTRL+F10", "CTRL+F11", "CTRL+F12",
    "CTRL+SHIFT+F1", "CTRL+SHIFT+F2", "CTRL+SHIFT+F3", "CTRL+SHIFT+F4", "CTRL+SHIFT+F5",
    "CTRL+SHIFT+F6", "CTRL+SHIFT+F7", "CTRL+SHIFT+F8", "CTRL+SHIFT+F9", "CTRL+SHIFT+F10",
    "CTRL+SHIFT+F11", "CTRL+SHIFT+F12", None, None, None, None, None, None, None, None, None, None,
    None, None, None, None, None, None, None, None, None, None, None, "CTRL+E", "CTRL+F", "CTRL+A",
    "CTRL+D", "CTRL+N", "CTRL+O", "SHIFT+DEL", "CTRL+INS", "SHIFT+INS", "ALT+BACKSPACE",
    "CTRL+PAGEUP", "PAGEUP",
    "PAGEDOWN", "CTRL+PAGEDOWN", "CTRL+G", "CTRL+R", "CTRL+P", "CTRL+B", "CTRL+K", "CTRL+T",
    "CTRL+Y",
    "CTRL+X", "CTRL+C", "CTRL+V", "SHIFT+F10", None, None, "CTRL+#"]


def assert_string_has_numeric(text: str, len_value: Optional[int] = None) -> bool:
    matched_value = re.search("\d+", text).group(0)
    if len_value is not None:
        if matched_value is not None:
            if len(matched_value) == len_value:
                return True
            else:
                return False
        else:
            return False
    else:
        if matched_value is not None:
            return True
        else:
            return False


def main_is_frozen() -> bool:
    return (hasattr(sys, "frozen") or # new py2exe
        hasattr(sys, "importers")) # old py2exe


def get_main_dir() -> str:
    if main_is_frozen():
        return os.path.dirname(sys.executable)
    elif hasattr(__builtins__,'__IPYTHON__'):
        return os.getcwd()
    else:
        return os.path.dirname(sys.argv[0])


def parent_func() -> str:
    return str(inspect.stack()[1].function)


def pad(value: str, length: int, char: Optional[str] = "0", right: Optional[bool] = False) -> str:
    if right:
        tmp = value.split(".")
        right_side = tmp[1]
        while len(right_side) < length:
            right_side = f"{right_side}{char}"
        value = f"{tmp[0]}.{right_side}"
    else:
        while len(value) < length:
            value = f"{char}{value}"
    return value


def string_generator(size: Optional[int]=10, chars: Optional[str]=string.ascii_uppercase + string.digits) -> str:
    selected_chars = []
    for i in range(size):
        selected_chars.append(random.choice(chars))
    return ''.join(selected_chars)


class Screenshot:
    def __init__(self) -> None:
        self.sct: MSSBase = mss()
        self.__directory: str = None
        self.__monitor: dict[str, int] = None
    
    @property
    def monitor(self) -> dict[str, int]:
        return self.__monitor
    
    @monitor.setter
    def monitor(self, value: int) -> None:
        self.__monitor = int(value)
    
    @property
    def screenshot_directory(self) -> str:
        return self.__directory
    
    @screenshot_directory.setter
    def screenshot_directory(self, value: str) -> None:
        __dir: str = os.path.join(get_main_dir(), value) 
        if not os.path.exists(__dir):
            try:
                os.mkdir(__dir)
            except Exception as err:
                raise FileNotFoundError(f"Directory {__dir} does not exist and was unable to be created automatically. Make sure you have the required access.")
        self.__directory = __dir
    
    def shot(self, monitor: Optional[int] = None, output: Optional[str] = None, name: Optional[str] = None, delay: Optional[float]= 2.0) -> list:
        if monitor:
            self.monitor(value=monitor)
        if output:
            self.screenshot_directory(value=output)
        else:
            if not self.__directory:
                self.screenshot_directory = "screenshots"
        time.sleep(delay)
        __name = f"{name}.jpg" if name is not None else f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
        return [x for x in self.sct.save(mon=self.__monitor, output=os.path.join(self.__directory, __name))]


class Timer:
    """
    A basic timer to user when waiting for element to be displayed. 
    """
    def __init__(self) -> None:
        self.start_time = time.time()

    def elapsed(self) -> float:
        return time.time() - self.start_time


class WindowHandler:
    def __init__(self, window_description: Optional[str] = None) -> None:
        self.window_handle_list: list = []
        self.window_handle: str = None
        self.window_description: str = window_description if window_description is not None else ""
        if len(self.window_handle_list) == 0:
            self.gather_window_list()
    
    def winEnumHandler(self, hwnd: str) -> None:
        if win32gui.IsWindowVisible(hwnd):
            self.window_handle_list.append((hwnd, win32gui.GetWindowText(hwnd)))
    
    def window_list_display(self, window_list: Optional[list] = None) -> str:
        _win_list = window_list if window_list is not None else self.window_handle_list
        import PySimpleGUI as sg
        layout = [[sg.Listbox(values=_win_list, size=(30, 6), enable_events=True, bind_return_key=True)]]
        window = sg.Window('Select Window', layout)
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Cancel':
                break
            else:
                window.close()
                if type(values) is list:
                    v: str = ""
                    if len(values) == 1:
                        v = str(values[0])
                    elif len(values) > 1:
                        v = "|".join(values)
                    else:
                        v = ""
                    return v
                elif type(values) is str:
                    return values
                else:
                    try:
                        return str(values)
                    except Exception:
                        return ""
    
    def gather_window_list(self) -> None:
        win32gui.EnumWindows(self.winEnumHandler, None)
    
    def close_window(self) -> None:
        self.window_handle = None
        for win in self.window_handle_list:
            if self.window_description == win[1]:
                self.window_handle = win[0]
            elif self.window_description in win[1]:
                self.window_handle = win[0] 
        if self.window_handle is not None:
            win32gui.PostMessage(self.window_handle, win32con.WM_CLOSE, 0, 0)
