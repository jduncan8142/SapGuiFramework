import win32com.client
import time
import re
import os
import sys
import datetime
import time
import string
import random
import logging
import base64
import atexit
from typing import Optional, Any
from mss import mss


def failed():
    return "FAIL"


def passed():
    return "PASS"


class SapLogger:
    def __init__(self, log_name: Optional[str] = None, log_path: Optional[str] = None, verbosity: Optional[int] = None) -> None:
        import log_conf as conf
        self.enabled: bool = conf.enable
        self.log_name: str = log_name if log_name is not None else conf.name
        self.log_path: str = log_path if log_path is not None else conf.path
        self.log_file: str = os.path.join(self.log_path, f"{self.log_name}.log")
        if not os.path.isdir(self.log_path):
            os.mkdir(self.log_path)
        if not os.path.isfile(self.log_file):
            with open(self.log_file, "w") as f:
                pass
        
        # Create custom logging level for screenshots
        SCREENSHOT_LEVELV_NUM = 60 
        logging.addLevelName(SCREENSHOT_LEVELV_NUM, "SHOT")
        def shot(self, message, *args, **kws):
            if self.isEnabledFor(SCREENSHOT_LEVELV_NUM):
                # Yes, logger takes its '*args' as 'args'.
                self._log(SCREENSHOT_LEVELV_NUM, message, args, **kws)
        logging.Logger.shot = shot

        # Create custom logging level for documentation
        DOUMENTATION_LEVELV_NUM = 70 
        logging.addLevelName(DOUMENTATION_LEVELV_NUM, "DOCUMENTATION")
        def documentation(self, message, *args, **kws):
            if self.isEnabledFor(DOUMENTATION_LEVELV_NUM):
                # Yes, logger takes its '*args' as 'args'.
                self._log(DOUMENTATION_LEVELV_NUM, message, args, **kws)
        logging.Logger.documentation = documentation

        self.log: logging.Logger = logging.getLogger(self.log_file)
        self.formatter: logging.Formatter = logging.Formatter(conf.format)
        self.file_handler: logging.FileHandler = logging.FileHandler(self.log_file, mode=conf.file_mode)
        self.file_handler.setFormatter(self.formatter)
        self.stream_handler: logging.StreamHandler = logging.StreamHandler()
        self.stream_handler.setFormatter(self.formatter)
        self.verbosity: int = verbosity if verbosity is not None else conf.verbosity
        match self.verbosity:
            case 5:
                self.log.setLevel(logging.DEBUG)
                self.file_handler.setLevel(logging.DEBUG)
                self.stream_handler.setLevel(logging.DEBUG)
            case 4:
                self.log.setLevel(logging.INFO)
                self.file_handler.setLevel(logging.INFO)
                self.stream_handler.setLevel(logging.NOTSET)
            case 3:
                self.log.setLevel(logging.WARNING)
                self.file_handler.setLevel(logging.WARNING)
                self.stream_handler.setLevel(logging.NOTSET)
            case 2:
                self.log.setLevel(logging.ERROR)
                self.file_handler.setLevel(logging.ERROR)
                self.stream_handler.setLevel(logging.NOTSET)
            case 1:
                self.log.setLevel(logging.CRITICAL)
                self.file_handler.setLevel(logging.CRITICAL)
                self.stream_handler.setLevel(logging.NOTSET)
            case _:
                self.log.setLevel(logging.NOTSET)
                self.file_handler.setLevel(logging.NOTSET)
                self.stream_handler.setLevel(logging.NOTSET)
        self.log.addHandler(self.file_handler)
        self.log.addHandler(self.stream_handler)


class Screenshot:
    def __init__(self) -> None:
        self.sct = mss()
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
        __dir: str = os.path.join(os.getcwd(), value) 
        if not os.path.exists(__dir):
            try:
                os.mkdir(__dir)
            except Exception as err:
                raise FileNotFoundError(f"Directory {__dir} does not exist and was unable to be created automatically. Make sure you have the required access.")
        self.__directory = __dir
    
    def shot(self, monitor: Optional[int] = None, output: Optional[str] = None, name: Optional[str] = None) -> list:
        if monitor:
            self.monitor(value=monitor)
        if output:
            self.screenshot_directory(value=output)
        else:
            if not self.__directory:
                self.screenshot_directory = "output"
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


class Gui:
    """
     Python Framework library for controlling the SAP GUI Desktop and focused 
     on testing business processes. The library uses the native SAP GUI scripting engine 
     for interaction with the desktop client application.
    """

    __version__ = '0.0.7'

    def __init__(
        self, 
        test_case: Optional[str] = "Default Test Case",
        log_path: Optional[str] = "output",
        verbosity: Optional[int] = 2,  
        screenshot_dir: Optional[str] = "output", 
        monitor: Optional[int] = 1, 
        explicit_wait: Optional[float] = 0.0, 
        connection_number: Optional[int] = 0, 
        session_number: Optional[int] = 0, 
        connection_name: Optional[str] = None, 
        date_format: Optional[str] = "%m/%d/%Y") -> None:
        atexit.register(self.cleanup)
        self.test_case_name: str = test_case
        self.subrc: int = 0
        self.logger = SapLogger(log_name=self.test_case_name, log_path=log_path, verbosity=verbosity)
        self.__connection_number: int = connection_number
        self.__session_number: int = session_number
        self.explicit_wait = explicit_wait
        self.connection_name: str = connection_name if connection_name is not None else ""
        self.sap_gui: win32com.client.CDispatch = None
        self.sap_app: win32com.client.CDispatch = None
        self.connection: win32com.client.CDispatch = None
        self.session: win32com.client.CDispatch = None
        self.screenshot: Screenshot = Screenshot()
        self.date_format = date_format

        if not os.path.exists(screenshot_dir):
            self.logger.log.debug(f"Screenshot directory {screenshot_dir} does not exist, creating it.")
            try:
                os.makedirs(screenshot_dir)
            except Exception as err:
                self.logger.log.error(f"Unable to create screenshot directory {screenshot_dir}")
        self.screenshot.screenshot_directory = screenshot_dir
        self.screenshot.monitor = monitor

        self.window: int = 0
        self.transaction: str = None
        self.sbar: win32com.client.CDispatch = None
        self.session_info: win32com.client.CDispatch = None
        self.text_elements = (
            "GuiTextField", 
            "GuiCTextField", 
            "GuiPasswordField", 
            "GuiLabel", 
            "GuiTitlebar", 
            "GuiStatusbar", 
            "GuiButton", 
            "GuiTab", 
            "GuiShell", 
            "GuiStatusPane"
            )
        self.task_status: str = None
        self.test_status: str = None
        self.test_case_failed: bool = False
        self.failed_tasks: list = []
        self.task: str = ""
    
    def transaction_does_not_exist_strings(self) -> tuple:
        return (
            f"Transactie {self.transaction} bestaat niet", 
            f"Transaction {self.transaction} does not exist", 
            f"Transaktion {self.transaction} existiert nicht"
            )
    
    def cleanup(self):
        if self.test_status is None:
            if self.test_case_failed or self.test_status == failed() or self.failed_tasks:
                self.test_status = failed()
            elif not self.test_case_failed and self.test_status != failed() and not self.failed_tasks:
                self.test_status = passed()
            else:
                self.test_status = "UNKNOWN > Check the logs."
        self.documentation(f"{self.test_case_name} completed with status: {self.test_status}")
        if self.failed_tasks:
            self.documentation(str("The following tasks failed: \n" + "\n".join([str(x) for x in self.failed_tasks]) + "\n"))

    def documentation(self, msg: str) -> None:
        self.task = msg
        self.logger.log.documentation(self.task)
    
    def fail(self, exit_on_error: Optional[bool] = True) -> None:
        self.task_status = failed()
        self.failed_tasks.append(self.task)
        self.test_status = failed()
        self.test_case_failed = True
        if exit:
            sys.exit()
    
    def task_passed(self) -> None:
        self.task_status = passed()

    def is_error(self) -> bool:
        if self.subrc != 0:
            return True
        else:
            return False

    def is_element(self, element: str) -> bool:
        try:
            self.session.findById(element)
            return True
        except:
            return False
    
    def log(self, msg: Any) -> None:
        self.logger.log.info(str(msg))

    def take_screenshot(self, screenshot_name: Optional[str] = "", msg: Optional[str] = "", logger: Optional[SapLogger] = None) -> None:
        _log = self.logger
        _file_names = None
        if logger:
            _log = logger
        if not screenshot_name:
            _file_names = self.screenshot.shot()
        else:
            _file_names = self.screenshot.shot(name=screenshot_name)
        if _file_names:
            for f in _file_names:
                encoded_img = None
                with open(f, "rb") as f_img:
                    encoded_img = base64.b64encode(f_img.read())
                _log.log.shot(f"{msg}|{f}|{encoded_img}")
            
    
    def wait(self, value: Optional[float] = None) -> None:
        if not value:
            time.sleep(self.explicit_wait)
        else:
            time.sleep(value)
    
    def get_element_type(self, id: str) -> str | None:
        try:
            return self.session.findById(id).type
        except Exception as err:
            self.logger.log.error(f"Unknown element id: {id} -> {err}")
            return None
    
    def pad(self, value: str, length: int, char: Optional[str] = "0", right: Optional[bool] = False) -> str:
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

    def connect_to_session(self) -> None:
        try:
            self.sap_gui = win32com.client.GetObject("SAPGUI")
            if not type(self.sap_gui) == win32com.client.CDispatch:
                self.logger.log.error(f"Error while getting SAP GUI object using win32com.client")
                return
            self.sap_app = self.sap_gui.GetScriptingEngine
            if not type(self.sap_app) == win32com.client.CDispatch:
                self.logger.log.error(f"Error while getting SAP scripting engine")
                self.sap_gui = None
                return
            self.connection = self.sap_app.Children(self.__connection_number)
            if not type(self.connection) == win32com.client.CDispatch:
                self.logger.log.error(f"Error while getting SAP connection to Window {self.__connection_number}")
                self.sap_app = None
                self.sap_gui = None
                return
            if self.connection.DisabledByServer == True:
                self.logger.log.error(f"SAP scripting is disable for this server")
                self.sap_app = None
                self.sap_gui = None
                return
            self.session = self.connection.Children(self.__session_number)
            if not type(self.session) == win32com.client.CDispatch:
                self.logger.log.error(f"Error while getting SAP session to Window {self.__session_number}")
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                return
            if self.session.Info.IsLowSpeedConnection == True:
                self.logger.log.error(f"SAP connect is listed as low speed, scripting not possible")
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                return
            self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.window}]/sbar")
            if not type(self.sbar) == win32com.client.CDispatch:
                self.logger.log.error(f"Unable to get status bar during session connection")
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                self.session = None
                return
            self.session_info = self.session.info
            self.task_passed()
        except:
            self.logger.log.error(f"Unknown error while establishing connection with SAP GUI -> {sys.exc_info()[0]}")
            self.fail()
        finally:
            self.sap_gui = None
            self.sap_app = None
            self.connection = None
            self.session = None
    
    def connect_to_existing_connection(self, connection_name: Optional[str] = None) -> None:
        if connection_name:
            self.connection_name = connection_name
        self.connection = self.sap_gui.Children(self.__connection_number)
        if self.connection.Description == self.connection_name:
            self.session = self.connection.children(self.session_number)
            self.wait()
            self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.session_number}]/wnd[{self.window}]/sbar")
            self.session_info = self.session.info
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="connect_to_existing_connection_error.jpg")
            self.logger.log.error(f"No existing connection for {self.connection_name} found.")
            self.fail()
    
    def open_connection(self, connection_name: Optional[str] = None):
        if not hasattr(self.sap_app, "OpenConnection"):
            try:
                self.sap_gui = win32com.client.GetObject("SAPGUI")
                if not type(self.sap_gui) == win32com.client.CDispatch:
                    self.logger.log.error(f"Error while getting SAP GUI object using win32com.client")
                    self.fail()
                self.sap_app = self.sap_gui.GetScriptingEngine
                if not type(self.sap_app) == win32com.client.CDispatch:
                    self.logger.log.error(f"Error while getting SAP scripting engine")
                    self.sap_gui = None
                    self.fail()
            except:
               self.logger.log.warning("SAP Login Pad not running")
               self.fail()
        if connection_name:
            self.connection_name = connection_name
        try:
            self.connection = self.sap_app.OpenConnection(self.connection_name, True)
        except Exception as err:
            self.logger.log.error(f"Cannot open connection {self.connection_name}, please check connection name -> {err}")
            self.fail()
        self.session = self.connection.children(self.__session_number)
        self.wait()
        self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.window}]/sbar")
        self.session_info = self.session.info
        self.task_passed()
    
    def exit(self) -> None:
        self.connection.closeSession(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]")
        self.connection.closeConnection()
        self.task_passed()
    
    def maximize_window(self) -> None:
        self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.window}]").maximize()
        self.task_passed()
    
    def restart_session(self, connection_name: str, delay: Optional[float] = 0.0) -> None:
        self.connection_name = connection_name if connection_name is not None else self.connection_name
        self.exit()
        self.open_connection(connection_name=self.connection_name)
        self.maximize_window()
        self.wait(value=delay)
        self.task_passed()
    
    def wait_for_element(self, id: str, timeout: Optional[float] = 60.0, exit_on_error: Optional[bool] = True) -> None:
        t = Timer()
        while not self.is_element(element=id) and t.elapsed() <= timeout:
            self.wait(value=1.0)
        if not self.is_element(element=id):
            self.take_screenshot(screenshot_name="wait_for_element_error.jpg")
            self.logger.log.error(f"Wait For Element could not find element with id {id}")
            self.fail(exit_on_error)
        self.task_passed()
    
    def get_statusbar_if_error(self, exit_on_error: Optional[bool] = True) -> str:
        try:
            if self.sbar.messageType == "E":
                return f"{self.sbar.findById('pane[0]').text} -> Message no. {self.sbar.messageId.strip('')}:{self.sbar.messageNumber}"
            else:
                return ""
        except:
            self.take_screenshot(screenshot_name="get_statusbar_if_error_error.jpg")
            self.logger.log.error(f"Error while checking if statusbar had error msg.")
            self.fail(exit_on_error)
    
    def get_status_msg(self) -> dict:
        try:
            msg_id = self.sbar.messageId
        except:
            msg_id = ""
        try:
            msg_number = self.sbar.messageNumber
        except:
            msg_number = ""
        try:
            msg_type = self.sbar.messageType
        except: 
            msg_type = ""
        try:
            msg = self.sbar.message
        except:
            msg = ""
        try:
            txt = self.sbar.text
        except:
            txt = ""
        return {"messageId": msg_id, "messageNumber": msg_number, "messageType": msg_type, "message": msg, "text": txt}
    
    def start_transaction(self, transaction: str) -> None:
        self.transaction = transaction.upper()
        self.session.startTransaction(self.transaction)
        self.wait()
        if (s_msg := str(self.sbar.findById('pane[0]').text).strip(" \n\r\t")) in self.transaction_does_not_exist_strings():
            self.take_screenshot(screenshot_name="start_transaction_error")
            self.logger.log.error(f"ValueError > {s_msg}")
            self.fail()
        else:
            self.logger.log.info(f"Started transaction {self.transaction} successfully > {s_msg}")
            self.task_passed()
    
    start = start_transaction
    
    def end_transaction(self) -> None:
        self.session.endTransaction()
        self.task_passed()
    
    end = end_transaction
    
    def send_command(self, command: str) -> None:
        try:
            self.session.sendCommand(command)
            self.task_passed()
        except Exception as err:
            self.take_screenshot(screenshot_name="send_command_error.jpg")
            self.logger.log.error(f"Error sending command {command} -> {err}")
            self.fail()

    def click_element(self, id: str = None) -> None:
        if (element_type := self.get_element_type(id)) in ("GuiTab", "GuiMenu"):
            self.session.findById(id).select()
        elif element_type == "GuiButton":
            self.session.findById(id).press()
        else:
            self.take_screenshot(screenshot_name="click_element_error.jpg")
            self.logger.log.error(f"You cannot use 'Click Element' on element id type {id}")
            self.fail()
        self.wait()
        self.task_passed()
    
    click = click_element

    def click_toolbar_button(self, table_id: str, button_id: str) -> None:
        self.element_should_be_present(table_id)
        try:
            self.session.findById(table_id).pressToolbarButton(button_id)
        except AttributeError:
            self.session.findById(table_id).pressButton(button_id)
        except Exception as err:
            self.take_screenshot(screenshot_name="click_toolbar_button_error.jpg")
            self.logger.log.error(f"Cannot find Table ID/Button ID: {' / '.join([table_id, button_id])}  <-->  {err}")
            self.fail()
        self.wait()
        self.task_passed()

    def doubleclick(self, id: str, item_id: str, column_id: str) -> None:
        if (element_type := self.get_element_type(id)) == "GuiShell":
            self.session.findById(id).doubleClickItem(item_id, column_id)
        else:
            self.take_screenshot(screenshot_name="doubleclick_element_error.jpg")
            self.logger.log.error(f"You cannot use 'doubleclick element' on element type {element_type}")
            self.fail()
        self.wait()
        self.task_passed()

    def assert_element_present(self, id: str, message: Optional[str] = None) -> None:
        if not self.is_element(element=id):
            self.take_screenshot(screenshot_name="assert_element_present_error.jpg")
            self.logger.log.error(message if message is not None else f"Cannot find element {id}")
            self.fail()
        self.task_passed()

    def assert_element_value(self, id: str, expected_value: str, message: Optional[str] = None) -> None:
        if self.is_element(element=id):
            actual_value = self.get_value(id=id)
            self.session.findById(id).setfocus()
            self.wait()
            self.task_passed()
        if (element_type := self.get_element_type(id)) in self.text_elements:
            if expected_value != actual_value:
                message = message if message is not None else f"Element value of {id} should be {expected_value}, but was {actual_value}"
                self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                self.logger.log.error(f"AssertEqualError > Element value of {id} should be {expected_value}, but was {actual_value}")
                self.fail()
        elif element_type in ("GuiCheckBox", "GuiRadioButton"):
            if expected_value := bool(expected_value):
                if not actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                    self.logger.log.error(f"AssertEqualError > Element value of {id} should be {expected_value}, but was {actual_value}")
                    self.fail()
            elif not expected_value:
                if actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                    self.logger.log.error(f"AssertEqualError > Element value of {id} should be {expected_value}, but was {actual_value}")
                    self.fail()
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
            self.logger.log.error(f"AssertEqualError > Element value of {id} should be {expected_value}, but was {actual_value}")
            self.fail()
    
    assert_element_value_equal = assert_element_value

    def assert_element_value_not_equal(self, id: str, expected_value: str, message: Optional[str] = None) -> None:
        if self.is_element(element=id):
            actual_value = self.get_value(id=id)
            self.session.findById(id).setfocus()
            self.wait()
            self.task_passed()
        if (element_type := self.get_element_type(id)) in self.text_elements:
            if expected_value == actual_value:
                message = message if message is not None else f"Element value of {id} should not be equal to {expected_value}"
                self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                self.fail()
        elif element_type in ("GuiCheckBox", "GuiRadioButton"):
            if expected_value := bool(expected_value):
                if not actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                    self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                    self.fail()
            elif not expected_value:
                if actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                    self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                    self.fail()
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
            self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
            self.fail()

    def assert_element_value_contains(self, id: str, expected_value: str, message: Optional[str] = None) -> None:
        if self.is_element(element=id):
            actual_value = self.get_value(id=id)
            self.session.findById(id).setfocus()
            self.wait()
            self.task_passed()
        if (element_type := self.get_element_type(id)) in self.text_elements:
            if expected_value != actual_value:
                message = message if message is not None else f"Element value of {id} does not contain {expected_value} but was {actual_value}"
                self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                self.logger.log.error(f"AssertContainsError > {message}")
                self.fail()
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error")
            self.logger.log.error(f"AssertContainsError > Element value of {id} does not contain {expected_value}, but was {actual_value}")
            self.fail()
        

    def get_cell_value(self, table_id: str, row_num: int, col_id: str) -> str | None:
        if self.is_element(element=table_id):
            try:
                _value = self.session.findById(table_id).getCellValue(row_num, col_id)
                self.task_passed()
                return _value
            except Exception as err:
                self.take_screenshot(screenshot_name="get_cell_value_error")
                self.logger.log.error(f"Cannot find cell value for table: {table_id}, row: {row_num}, and column: {col_id} -> {err}")
                self.fail()

    def set_combobox(self, id: str, key: str) -> None:
        if (element_type := self.get_element_type(id)) == "GuiComboBox":
            self.session.findById(id).key = key
            self.logger.log.info(f"ComboBox value {key} selected from {id}")
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="set_combobox_error")
            self.logger.log.error(f"Element type {element_type} for element {id} has no set key method.")
            self.fail()
    
    combobox = set_combobox

    def get_element_location(self, id: str) -> tuple[int] | None:
        _location = (self.session.findById(id).screenLeft, self.session.findById(id).screenTop) if self.is_element(element=id) else None
        if _location:
            self.task_passed()
        else:
            self.fail()

    def get_element_type(self, id) -> Any:
        try:
            _type = self.session.findById(id).type
            self.task_passed()
            return _type
        except Exception as err:
            self.take_screenshot(screenshot_name="get_element_type_error")
            self.logger.log.error(f"Cannot find element type for id: {id} -> {err}")
            self.fail()

    def get_row_count(self, table_id) -> int:
        try:
            _count = self.session.findById(table_id).rowCount if self.is_element(element=table_id) else 0
            self.task_passed()
            return _count
        except Exception as err:
            self.take_screenshot(screenshot_name="get_row_count_error")
            self.logger.log.error(f"Cannot find row count for table: {table_id} -> {err}")
            self.fail()

    def get_scroll_position(self, id: str) -> int:
        self.wait()
        try:
            _position = int(self.session.findById(id).verticalScrollbar.position) if self.is_element(element=id) else 0
            self.task_passed()
            return _position
        except Exception as err:
            self.take_screenshot(screenshot_name="get_scroll_position_error")
            self.logger.log.error(f"Cannot get scrollbar position for: {id} -> {err}")
            self.fail()

    def get_window_title(self, id: str) -> str:
        try:
            _title =  self.session.findById(id).text if self.is_element(element=id) else ""
            self.task_passed()
            return _title
        except Exception as err:
            self.take_screenshot(screenshot_name="get_window_title_error")
            self.logger.log.error(f"Cannot find window with locator {id} -> {err}")
            self.fail()

    def get_value(self, id: str) -> str | bool:
        try:
            _value = None
            if (element_type := self.get_element_type(id)) in self.text_elements:
                _value = self.session.findById(id).text
            elif element_type in ("GuiCheckBox", "GuiRadioButton"):
                _value = self.session.findById(id).selected
            elif element_type == "GuiComboBox":
                _value = str(self.session.findById(id).text).strip()
            else:
                self.take_screenshot(screenshot_name="get_value_warning")
                self.logger.log.error(f"Cannot get value for element type {element_type} for id {id}")
            if _value:
                self.task_passed()
            else:
                self.fail()
        except Exception as err:
            self.take_screenshot(screenshot_name="get_value_error.jpg")
            self.logger.log.error(f"Cannot get value for element type {element_type} for id {id} -> {err}")
            self.fail()

    def input_text(self, id: str, text: str) -> None:
        if (element_type := self.get_element_type(id)) in self.text_elements:
            self.session.findById(id).text = text
            if element_type != "GuiPasswordField":
                self.logger.log.info(f"Input {text} into text field {id}")
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="input_text_error.jpg")
            self.logger.log.error(f"Cannot use keyword 'input text' for element type {element_type}")
            self.fail()
    
    text = input_text

    def string_generator(size: int = 6, chars: str = string.ascii_uppercase + string.digits):
        return ''.join(random.choice(chars) for _ in range(size))

    def input_random_value(self, id: str, text: str, prefix: bool = False, suffix: bool = False, date_time: bool = False, random: bool = False) -> str:
        dt: str = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f") if date_time else ""
        rs: str = self.string_generator() if random else ""
        tmp: str = text
        if prefix:
            tmp = f"{dt}_{rs}{tmp}"
        if suffix:
            tmp = f"{tmp}_{dt}{rs}"
        self.input_text(id=id, text=tmp)
        return tmp
    
    def input_current_date(self, id: str, format: Optional[str] = "%m/%d/%Y") -> None:
        self.input_text(id=id, text=datetime.datetime.now().strftime(format))

    def maximize_window(self, window: Optional[int] = None) -> None:
        if window:
            self.window = window
        try:
            self.session.findById(f"wnd[{self.window}]").maximize()
            self.wait()
        except Exception as err:
            self.take_screenshot(screenshot_name="maximize_window")
            self.logger.log.error(f"Cannot maximize window wnd[{self.window}] -> {err}")

    def set_vertical_scroll(self, id: str, position: int) -> None:
        if self.is_element(id):
            self.session.findById(id).verticalScrollbar.position = position
            self.wait()
            self.task_passed()
        else:
            self.fail()

    def set_horizontal_scroll(self, id: str, position: int) -> None:
        if self.is_element(id):
            self.session.findById(id).horizontalScrollbar.position = position
            self.wait()
            self.task_passed()
        else:
            self.fail()

    def get_vertical_scroll(self, id: str) -> int | None:
        try:
            _vs = self.session.findById(id).verticalScrollbar.position if self.is_element(id) else None
            self.task_passed()
            return _vs
        except Exception as err:
            self.take_screenshot(screenshot_name="get_vertical_scroll")
            self.logger.log.error(f"Cannot get vertical scroll position -> {err}")
            self.fail()

    def get_horizontal_scroll(self, id: str) -> int | None:
        try:
            _hs = self.session.findById(id).horizontalScrollbar.position if self.is_element(id) else None
            self.task_passed()
            return _hs
        except Exception as err:
            self.take_screenshot(screenshot_name="get_horizontal_scroll")
            self.logger.log.error(f"Cannot get horizontal scroll position -> {err}")
            self.fail()

    def select_checkbox(self, id: str) -> None:
        if (element_type := self.get_element_type(id)) == "GuiCheckBox":
            self.session.findById(id).selected = True
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="select_checkbox_error")
            self.logger.log.error(f"Cannot use keyword 'select checkbox' for element type {element_type}")
            self.fail()

    def unselect_checkbox(self, id: str) -> None:
        if (element_type := self.get_element_type(id)) == "GuiCheckBox":
            self.session.findById(id).selected = False
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="select_checkbox_error")
            self.logger.log.error(f"Cannot use keyword 'unselect checkbox' for element type {element_type}")
            self.fail()

    def set_cell_value(self, table_id, row_num, col_id, text):
        if self.is_element(element=table_id):
            try:
                self.session.findById(table_id).modifyCell(row_num, col_id, text)
                self.logger.log.info(f"Input {text} into cell ({row_num}, {col_id})")
                self.wait()
                self.task_passed()
            except Exception as err:
                self.take_screenshot(screenshot_name="set_cell_value_error.jpg")
                self.logger.log.error(f"Failed entering {text} into cell ({row_num}, {col_id}) -> {err}")

    def send_vkey(self, vkey: str, window: int = 0) -> None:
        vkey_id = str(vkey)
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
        if not vkey_id.isdigit():
            search_comb = vkey_id.upper()
            search_comb = search_comb.replace(" ", "")
            search_comb = search_comb.replace("CONTROL", "CTRL")
            search_comb = search_comb.replace("DELETE", "DEL")
            search_comb = search_comb.replace("INSERT", "INS")
            try:
                vkey_id = vkeys.index(search_comb)
            except ValueError:
                if search_comb == "CTRL+S":
                    vkey_id = 11
                elif search_comb == "ESC":
                    vkey_id = 12
                else:
                    self.logger.log.error(f"Cannot find given Vkey {vkey}, provide a valid Vkey number or combination")
                    self.fail()
        try:
            self.session.findById(f"wnd[{self.window}]").sendVKey(vkey_id)
            self.wait()
            self.task_passed()
        except Exception as err:
            self.take_screenshot(screenshot_name="send_vkey_error")
            self.logger.log.error(f"Cannot send Vkey to window wnd[{self.window}]]")
            self.fail()

    def select_context_menu_item(self, id: str, menu_id: str, item_id: str) -> None:
        if self.is_element(element=id):
            if hasattr(self.session.findById(id), "nodeContextMenu"):
                self.session.findById(id).nodeContextMenu(menu_id)
                self.session.findById(id).selectContextMenuItem(item_id)
                self.wait()
                self.task_passed()
            elif hasattr(self.session.findById(id), "pressContextButton"):
                self.session.findById(id).pressContextButton(menu_id)
                self.session.findById(id).selectContextMenuItem(item_id)
                self.wait()
                self.task_passed()
            else:
                self.take_screenshot(screenshot_name="select_context_menu_item_error")
                self.logger.log.error(f"Cannot use keyword 'Select Context Menu Item' with element type {self.get_element_type(id)}")
                self.fail()

    def select_from_list_by_label(self, id: str, value: str) -> None:
        if (element_type := self.get_element_type(id)) == "GuiComboBox":
            self.session.findById(id).key = value
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="select_from_list_by_label_error")
            self.logger.log.error(f"Cannot use keyword Select From List By Label with element type {element_type}")
            self.fail()

    def select_node(self, tree_id: str, node_id: str, expand: bool = False):
        if self.is_element(element=tree_id):
            self.session.findById(tree_id).selectedNode = node_id
            if expand:
                try:
                    self.session.findById(tree_id).expandNode(node_id)
                except:
                    self.take_screenshot(screenshot_name="expand_node")
                    self.logger.log.error(f"Unable to expand node {node_id} from tree {tree_id}")
                    self.fail()
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="select_node")
            self.logger.log.error(f"Unable to select node {node_id} from tree {tree_id}")
            self.fail()

    def select_node_link(self, tree_id: str, link_id1: str, link_id2: str) -> None:
        if self.is_element(element=tree_id):
            self.session.findById(tree_id).selectItem(link_id1, link_id2)
            self.session.findById(tree_id).clickLink(link_id1, link_id2)
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="select_node_link")
            self.logger.log.error(f"Unable to select node {link_id1} and click link {link_id2} from tree {tree_id}")
            self.fail()

    def select_radio_button(self, id: str) -> None:
        if (element_type := self.get_element_type(id)) == "GuiRadioButton":
            self.session.findById(id).selected = True
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="select_radio_button_error")
            self.logger.log.error(f"Cannot use keyword Select Radio Button with element type {element_type}")
            self.fail()
        

    def select_table_column(self, table_id: str, column_id: str) -> None:
        if self.is_element(element=table_id):
            try:
                self.session.findById(table_id).selectColumn(column_id)
                self.wait()
                self.task_passed()
            except Exception as err:
                self.take_screenshot(screenshot_name="select_table_column_error")
                self.logger.log.error(f"Cannot find column ID: {column_id} for table {table_id}")
                self.fail()
            

    def select_table_row(self, table_id: str, row_num: int) -> None:
        if (element_type := self.get_element_type(table_id)) == "GuiTableControl":
            id = self.session.findById(table_id).getAbsoluteRow(row_num)
            id.selected = -1
            self.wait()
            self.task_passed()
        else:
            try:
                self.session.findById(table_id).selectedRows = row_num
                self.wait()
                self.task_passed()
            except Exception as err:
                self.take_screenshot(screenshot_name="select_table_row_error")
                self.logger.log.error(f"Cannot use keyword Select Table Row for element type {element_type} -> {err}")
                self.fail()
        
    
    def try_and_continue(self, func_name: str, *args, **kwargs) -> Any:
        result = None
        self.wait(1.0)
        try:
            if hasattr(self, func_name) and callable(func := getattr(self, func_name)):
                result = func(*args, **kwargs)
        except Exception:
            pass
        return result
    
    def get_next_empty_table_row(self, table_id: str, column_index: Optional[int] = 0) -> None:
        try:
            table = self.session.findById(table_id)
            rows = table.rows
            for i in range(rows.count):
                row = rows.elementAt(i)
                if row.elementAt(column_index).text == "":
                    self.task_passed()
                    return i
            self.fail()
        except Exception as err:
            self.take_screenshot(screenshot_name="get_next_empty_table_row")
            self.logger.log.error(f"Cannot get next empty table row for table {table_id} -> {err}")
            self.fail()
    
    def insert_in_table(self, table_id: str, value: str, column_index: int = 0, row_index: Optional[int] = None) -> None:
        if not row_index:
            row_index = self.get_next_empty_table_row(table_id=table_id, column_index=column_index)
        table = self.session.findById(table_id)
        cell = table.getCell(row_index, column_index)
        (element_type := cell.type)
        if (element_type := cell.type) == "GuiComboBox":
            cell.key = value
            self.wait()
            self.task_passed()
        elif element_type == "GuiCTextField":
            cell.text = value
            self.wait()
            self.task_passed()
        else:
            self.take_screenshot(screenshot_name="insert_in_table")
            self.logger.log.error(f"Cannot inset {value} in table {table_id}")
            self.fail()
    
    def enter(self) -> None:
        self.send_vkey(vkey="ENTER")
    
    def save(self) -> None:
        self.send_vkey(vkey="CTRL+S")


class SalesOrder:
    def __init__(self, sap: Gui) -> None:
        self.sap: Gui = sap
        self.new_sales_order: str = None
        self.status_msg: str = None
        self.today = datetime.datetime.now().strftime(self.sap.date_format)
    
    def set_new_sales_order(self, msg: str) -> None:
        self.new_sales_order = re.search("\d+", msg).group(0)
    
    def va01(self) -> None:
        self.sap.start_transaction(transaction="VA01")
    
    def va01_initial_screen(self, order_type: str, sales_org: str, dist_ch: str, division: str, sales_office: Optional[str] = "", sales_group: Optional[str] = "", 
        press_enter: Optional[bool] = True) -> None:
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-AUART", text=order_type)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VKORG", text=sales_org)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VTWEG", text=dist_ch)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-SPART", text=division)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VKBUR", text=sales_office)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VKGRP", text=sales_group)
        if press_enter:
            self.sap.send_vkey(vkey="Enter")

    def va01_header(self, sold_to: str, ship_to: str, cust_ref: Optional[str] = None, cust_ref_date: Optional[str] = None, press_enter: Optional[bool] = True) -> None:
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR", text=sold_to)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR", text=ship_to)
        if cust_ref:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=cust_ref)
        else:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=datetime.datetime.now().strftime("%Y%m%d%H%M%S%f"))
        if cust_ref_date:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK", text=cust_ref_date)
        if press_enter:
            self.sap.send_vkey(vkey="Enter")
        # Handle status msg about duplicate PO values
        result = self.sap.get_status_msg()
        if result["messageId"].strip(" \n\r\t") == "V4" and result["messageNumber"] == "115" and result["messageType"] == "W":
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=datetime.datetime.now().strftime("%Y%m%d%H%M%S%f"))
            self.sap.enter()
    
    def va01_sales_tab(self, req_del_date_format: Optional[str] = None, req_del_date: Optional[str] = None, delver_plant: Optional[str] = None, delivery_block: Optional[str] = None, 
        billing_block: Optional[str] = None, pricing_date: Optional[str] = None, pyt_terms: Optional[str] = None, inco_version: Optional[str] = None, incoterms: Optional[str] = None, 
        inco_location1: Optional[str] = None, order_reason: Optional[str] = None, press_enter: Optional[bool] = True, complete_dlv: Optional[bool] = False) -> None:
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01")
        if req_del_date_format:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KPRGBZ", text=req_del_date_format)
        if req_del_date:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT", text=req_del_date)
        if delver_plant:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-DWERK", text=delver_plant)
        if complete_dlv:
            self.sap.select_checkbox(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/chkVBAK-AUTLF")
        else:
            self.sap.unselect_checkbox(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/chkVBAK-AUTLF")
        if delivery_block:
            self.sap.set_combobox(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK", key=delivery_block)
        if billing_block:
            self.sap.set_combobox(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK", key=billing_block)
        if pyt_terms:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBKD-ZTERM", text=pyt_terms)
        if pricing_date:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBKD-PRSDT", text=pricing_date)
        if inco_version:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBKD-INCOV", text=inco_version)
        if incoterms:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBKD-INCO1", text=incoterms)
        if inco_location1:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBKD-INCO2_L", text=inco_location1)
        if order_reason:
            self.set_combobox(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU", key=order_reason)
        if press_enter:
            self.sap.send_vkey(vkey="Enter")
    
    def va01_line_items(self, line_items: list[dict], press_enter: Optional[bool] = True) -> None:
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01")
        for item in line_items:
            self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POAN")
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,1]", text=item["material"])
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,1]", text=item["target_quantity"])
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-VRKME[3,1]", text=item["uom"])
            if "customer_material" in item:
                self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-KDMAT[6,1]", text=item["customer_material"])
            if "item_cat" in item:
                self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PSTYV[7,1]", text=item["item_cat"])
            if press_enter:
                self.sap.send_vkey(vkey="Enter")
    
    def create_new_sales_order(self, data: object, transaction: Optional[str] = "VA01", random_po: Optional[bool] = True) -> None:
        self.sap.start_transaction(transaction=transaction)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-AUART", text=data.order_type)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VKORG", text=data.sales_org)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-VTWEG", text=data.dist_ch)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtVBAK-SPART", text=data.division)
        self.sap.send_vkey(vkey="Enter")
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR", text=data.sold_to)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR", text=data.ship_to)
        if random_po:
            po = self.sap.input_random_value(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=data.order_type, suffix=True, date_time=True)
        else:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=data.po)
        self.sap.send_vkey(vkey="Enter")
        for item in data.line_items:
            self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POAN")
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,1]", text=item["material"])
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,1]", text=item["target_quantity"])
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-VRKME[3,1]", text=item["uom"])
            self.sap.send_vkey(vkey="Enter")
    
    def create_sales_order_from_reference(self, data: object, transaction: Optional[str] = "VA01") -> None:
        pass
    
    def va01_update_shipping_condition(self, shipping_condition: str, press_enter: Optional[bool] = True) -> None:
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\02")
        self.sap.set_combobox(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4302/cmbVBAK-VSBED", key=shipping_condition)
        if press_enter:
            self.sap.send_vkey(vkey="Enter")
        self.sap.try_and_continue(func_name="wait_for_element", id="/app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-VAROPTION1", exit_on_error=False)
        self.sap.try_and_continue(func_name="click_element", id="/app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-VAROPTION1", exit_on_error=False)
    
    def update_partners(self, partner_type: str, partner_number: str, press_enter: Optional[bool] = True) -> None:
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09")
        self.sap.insert_in_table(table_id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW", value=partner_type, column_index=0)
        self.sap.insert_in_table(table_id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW", value=partner_number, column_index=6)
        if press_enter:
            self.sap.send_vkey(vkey="Enter")

    def update_outputs(self, output_record: str, printer: str, print_immediate: Optional[bool] = True, press_enter: Optional[bool] = True) -> None:
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/mbar/menu[3]/menu[9]/menu[0]")
        self.sap.insert_in_table(table_id="/app/con[0]/ses[0]/wnd[0]/usr/tblSAPDV70ATC_NAST3", value=output_record, column_index=1)
        self.sap.send_vkey(vkey="Enter")
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[2]")
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtNAST-LDEST", text=printer)
        if print_immediate:
            self.sap.select_checkbox(id="/app/con[0]/ses[0]/wnd[0]/usr/chkNAST-DIMME")
        else:
            self.sap.unselect_checkbox(id="/app/con[0]/ses[0]/wnd[0]/usr/chkNAST-DIMME")
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]")
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]") 



