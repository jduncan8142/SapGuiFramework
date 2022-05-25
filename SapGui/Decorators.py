import functools
from Utilities.Utilities import *


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
