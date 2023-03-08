import datetime
import time
from mss import mss
from mss.base import MSSBase
from pathlib import Path
from typing import Optional
import random
import sys
import win32gui
import win32con
import re
import functools
import string
import inspect


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


def get_date_time_string(format: Optional[str] = "%m/%d/%Y%H%M%S") -> str:
    return datetime.datetime.now().strftime(format)
    

def assert_string_has_numeric(text: str, len_value: Optional[int] = None) -> bool:
    matched_value = re.search("\d+", text).group(0)
    if len_value is None:
        return False 
    if matched_value is None:
        return False
    if len(matched_value) != len_value:
        return False
    return True


def main_is_frozen() -> bool:
    return (hasattr(sys, "frozen") or # new py2exe
        hasattr(sys, "importers")) # old py2exe


def get_main_dir() -> Path:
    if main_is_frozen():
        return Path(sys.executable)
    elif hasattr(__builtins__,'__IPYTHON__'):
        return Path.cwd()
    else:
        return Path(*Path(sys.argv[0]).parts[:-4])
    

def explicit_wait_before(_func = None, *, wait_time: float = 0.0):
    def decorator_explicit_wait_before(func):
        @functools.wraps(func)
        def wait_wrapper(*args, **kwargs):
            time.sleep(wait_time)
            return func(*args, **kwargs)
        return wait_wrapper
    if _func is None:
        return decorator_explicit_wait_before
    else:
        return decorator_explicit_wait_before(_func)


def explicit_wait_after(_func = None, *, wait_time: float = 0.0):
    def decorator_explicit_wait_after(func):
        @functools.wraps(func)
        def wait_wrapper(*args, **kwargs):
            value = func(*args, **kwargs)
            time.sleep(wait_time)
            return value
        return wait_wrapper
    if _func is None:
        return decorator_explicit_wait_after
    else:
        return decorator_explicit_wait_after(_func)


class Timer:
    """
    A basic timer to use when waiting for element to be displayed. 
    """
    def __init__(self) -> None:
        self.start_time = time.time()

    def elapsed(self) -> float:
        return time.time() - self.start_time


class Screenshot:
    def __init__(self) -> None:
        self.sct: MSSBase = mss()
        self.__directory: Path = None
        self.__monitor: dict[str, int] = None
    
    @property
    def monitor(self) -> dict[str, int]:
        return self.__monitor
    
    @monitor.setter
    def monitor(self, value: int) -> None:
        self.__monitor = int(value)
    
    @property
    def screenshot_directory(self) -> Path:
        return self.__directory
    
    @screenshot_directory.setter
    def screenshot_directory(self, value: str) -> None:
        __path = get_main_dir()
        __dir: Path = __path.joinpath(value) 
        if not __dir.exists():
            try:
                __path.mkdir(__dir)
            except Exception as err:
                raise FileNotFoundError(f"Directory {__dir} does not exist and was unable to be created automatically. Make sure you have the required access. > {err}")
        self.__directory = __dir
    
    def shot(self, monitor: Optional[int] = None, output: Optional[str] = None, name: Optional[str] = None, delay: Optional[float]= 2.0) -> list:
        if monitor:
            self.monitor(value=monitor)
        if output:
            self.screenshot_directory(value=output)
        else:
            if not self.__directory:
                __path = get_main_dir()
                self.screenshot_directory = __path.joinpath("screenshots")
        time.sleep(delay)
        __name = f"{name}.jpg" if name is not None else f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
        return [x for x in self.sct.save(mon=self.__monitor, output=self.__directory.joinpath(__name))]
    

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

