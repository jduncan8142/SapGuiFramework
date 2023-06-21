import datetime
import time
from pathlib import Path
from typing import Optional
import random
import sys
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


def parse_sql_select(statement: str) -> list:
    statement = " ".join([x.strip('\t') for x in statement.strip("\n\r").upper().split(';')])   
    statement = statement + ' WHERE ' if 'WHERE' not in statement else statement

    regex = re.compile("SELECT(.*)FROM(.*)WHERE(.*)")

    parts = regex.findall(statement)
    parts = parts[0]
    select = [x.strip() for x in parts[0].split(',')]
    top = select[0].split(" ")
    if len(top) == 3:
        if "TOP" in top:
            select[0] = top[2]
            top = top[0:2]
        else:
            top = []
    else:
        top = []
    frm = parts[1].strip()
    where = parts[2].strip()

    # splits by spaces but ignores quoted string with ''
    PATTERN = re.compile(r"""((?:[^ '"]|'[^']*'|"[^"]*")+)""")
    where = PATTERN.split(where)[1::2]

    cleaned = [select, top, frm, where]
    return cleaned


class Timer:
    """
    A basic timer to use when waiting for element to be displayed. 
    """
    def __init__(self) -> None:
        self.start_time = time.time()

    def elapsed(self) -> float:
        return time.time() - self.start_time
