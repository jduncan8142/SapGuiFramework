import win32com.client
import time
from mss import mss
import os
import sys
import datetime
import time
import string
import random
from robot.api import logger
from typing import Optional


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
        __name = name if name is not None else f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
        return [x for x in self.sct.save(mon=self.__monitor, output=os.path.join(self.__directory, __name))]


class Timer:
    """
    A basic timer to user when waiting for element to be displayed. 
    """
    def __init__(self) -> None:
        self.start_time = time.time()

    def elapsed(self) -> float:
        return time.time() - self.start_time


class SapGuiRobot:
    """
     A Robocorp Robot Framework library for controlling the SAP GUI Desktop and focused 
     on testing business processes. The library uses the native SAP GUI scripting engine 
     for interaction with the desktop client application.
    """

    __version__ = '0.0.4'
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'

    def __init__(
        self, 
        screenshot_dir: Optional[str] = "output", 
        monitor: int = 1, 
        explicit_wait: Optional[float] = 0.0, 
        connection_number: Optional[int] = 0, 
        session_number: Optional[int] = 0, 
        connection_name: Optional[str] = None, 
        log_level: Optional[int] = 1) -> None:
        self.log_level: int = 1
        try:
            self.log_level = int(log_level)
        except Exception as err:
            if type(log_level) is not int:
                raise ValueError("log_level must be an int between 0 and 9")
            else:
                raise Exception(f"Unknown error while setting log_level -> {err}")

        self.subrc: int = 0
        self.__explicit_wait: float = explicit_wait
        self.__connection_number: int = connection_number
        self.__session_number: int = session_number
        self.connection_name: str = connection_name if connection_name is not None else ""
        self.sap_gui: win32com.client.CDispatch = None
        self.sap_app: win32com.client.CDispatch = None
        self.connection: win32com.client.CDispatch = None
        self.session: win32com.client.CDispatch = None
        self.screenshot: Screenshot = Screenshot()

        if not os.path.exists(screenshot_dir):
            logger.debug(f"Screenshot directory {screenshot_dir} does not exist, creating it.")
            try:
                os.makedirs(screenshot_dir)
            except Exception as err:
                logger.error(f"Unable to create screenshot directory {screenshot_dir}")
        self.screenshot.screenshot_directory = screenshot_dir
        self.screenshot.monitor = monitor

        self.window: int = 0
        self.transaction: str = None
        self.sbar: win32com.client.CDispatch = None
        self.session_info: win32com.client.CDispatch = None

        self.text_elements = ("GuiTextField", "GuiCTextField", "GuiPasswordField", "GuiLabel", "GuiTitlebar", "GuiStatusbar", "GuiButton", "GuiTab", "GuiShell", "GuiStatusPane")

    @property
    def explicit_wait(self) -> float:
        return self.__explicit_wait

    @explicit_wait.setter
    def explicit_wait(self, value: float = 0.0) -> None:
        try:
            self.explicit_wait = float(value)
            logger.debug(f"explicit_wait time set to: {value}")
        except TypeError:
            self.explicit_wait = float(0.0)
            logger.debug(f"Unable to set explicit_wait time with your given value. {value} cannot be converted to a float.")
        except Exception as err:
            self.explicit_wait = float(0.0)
            logger.debug(f"Unable to set explicit_wait time -> {err}")
    
    @property
    def connection_number(self) -> int:
        return self.__connection_number
    
    @connection_number.setter
    def connection_number(self, value: int = 0) -> None:
        try:
            self.__connection_number = int(value)
            logger.debug(f"connection_number set to: {value}")
        except TypeError:
            self.connection_number = int(0)
            logger.debug(f"Unable to set connection_number with your given value. {value} cannot be converted to a int.")
        except Exception as err:
            self.connection_number = int(0)
            logger.debug(f"Unable to set connection_number ->  {err}")
    
    @property
    def session_number(self) -> int:
        return self.__session_number

    @session_number.setter
    def session_number(self, value: int = 0) -> None:
        try:
            self.__session_number = int(value)
        except TypeError as err:
            logger.debug(err)
            self.session_number = int(0)

    def is_error(self) -> bool:
        if self.subrc != 0:
            return True
        else:
            return False

    def is_element(self, element: str) -> bool:
        try:
            self.session.findById(id)
            self.set_focus(id)
            return True
        except:
            return False

    def take_screenshot(self, screenshot_name: str = "") -> None:
        if not screenshot_name:
            self.screenshot.shot()
        else:
            self.screenshot.shot(name=screenshot_name)
    
    def wait(self, value: Optional[float] = None) -> None:
        if not value:
            time.sleep(self.explicit_wait)
        else:
            time.sleep(value)
    
    def get_element_type(self, id: str) -> str | None:
        try:
            return self.session.findById(id).type
        except Exception as err:
            logger.error(f"Unknown element id: {id} -> {err}")
            return None

    def connect_to_session(self) -> None:
        try:
            self.sap_gui = win32com.client.GetObject("SAPGUI")
            if not type(self.sap_gui) == win32com.client.CDispatch:
                logger.error(f"Error while getting SAP GUI object using win32com.client")
                return
            self.sap_app = self.sap_gui.GetScriptingEngine
            if not type(self.sap_app) == win32com.client.CDispatch:
                logger.error(f"Error while getting SAP scripting engine")
                self.sap_gui = None
                return
            self.connection = self.sap_app.Children(self.connection_number)
            if not type(self.connection) == win32com.client.CDispatch:
                logger.error(f"Error while getting SAP connection to Window {self.connection_number}")
                self.sap_app = None
                self.sap_gui = None
                return
            if self.connection.DisabledByServer == True:
                logger.error(f"SAP scripting is disable for this server")
                self.sap_app = None
                self.sap_gui = None
                return
            self.session = self.connection.Children(self.session_number)
            if not type(self.session) == win32com.client.CDispatch:
                logger.error(f"Error while getting SAP session to Window {self.session_number}")
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                return
            if self.session.Info.IsLowSpeedConnection == True:
                logger.error(f"SAP connect is listed as low speed, scripting not possible")
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                return
            self.sbar = self.session.findById(f"/app/con[{self.connection_number}]/ses[{self.session_number}]/wnd[{self.window}]/sbar")
            if not type(self.sbar) == win32com.client.CDispatch:
                logger.error(f"Unable to get status bar during session connection")
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                self.session = None
                return
            self.session_info = self.session.info
        except:
            logger.error(f"Unknown error while establishing connection with SAP GUI -> {sys.exc_info()[0]}")
        finally:
            self.sap_gui = None
            self.sap_app = None
            self.connection = None
            self.session = None
    
    def connect_to_existing_connection(self, connection_name: Optional[str] = None) -> None:
        if connection_name:
            self.connection_name = connection_name
        self.connection = self.sap_gui.Children(self.connection_number)
        if self.connection.Description == self.connection_name:
            self.session = self.connection.children(self.session_number)
            self.wait()
            self.sbar = self.session.findById(f"/app/con[{self.connection_number}]/ses[{self.session_number}]/wnd[{self.window}]/sbar")
            self.session_info = self.session.info
        else:
            self.take_screenshot(screenshot_name="connect_to_existing_connection_error.jpg")
            raise ValueError(f"No existing connection for {self.connection_name} found.")
    
    def open_connection(self, connection_name: Optional[str] = None):
        if not hasattr(self.sap_app, "OpenConnection"):
            try:
                self.sap_gui = win32com.client.GetObject("SAPGUI")
                if not type(self.sap_gui) == win32com.client.CDispatch:
                    logger.error(f"Error while getting SAP GUI object using win32com.client")
                    return
                self.sap_app = self.sap_gui.GetScriptingEngine
                if not type(self.sap_app) == win32com.client.CDispatch:
                    logger.error(f"Error while getting SAP scripting engine")
                    self.sap_gui = None
                    return
            except:
                raise Warning("SAP Login Pad not running")
        if connection_name:
            self.connection_name = connection_name
        try:
            self.connection = self.sap_app.OpenConnection(self.connection_name, True)
        except Exception as err:
            raise ValueError(f"Cannot open connection {self.connection_name}, please check connection name -> {err}")
        self.session = self.connection.children(self.session_number)
        self.wait()
        self.sbar = self.session.findById(f"/app/con[{self.connection_number}]/ses[{self.session_number}]/wnd[{self.window}]/sbar")
        self.session_info = self.session.info
    
    def wait_for_element(self, id: str, timeout: Optional[float] = 60.0) -> None:
        t = Timer()
        while not self.is_element(element=id) and t.elapsed() <= timeout:
            self.wait(value=1.0)
        if not self.is_element(element=id):
            self.take_screenshot(screenshot_name="wait_for_element_error.jpg")
            raise ValueError(f"Wait For Element could not find element with id {id}")
    
    def get_statusbar_if_error(self) -> str:
        try:
            if self.sbar.messageType == "E":
                return f"{self.sbar.findById('pane[0]').text} -> Message no. {self.sbar.messageId.strip('')}:{self.sbar.messageNumber}"
            else:
                return ""
        except:
            self.take_screenshot(screenshot_name="get_statusbar_if_error_error.jpg")
            raise ValueError(f"Error while checking if statusbar had error msg.")
    
    def start_transaction(self, transaction: str) -> None:
        if transaction:
            self.transaction = transaction.upper()
            self.session.startTransaction(self.transaction)
            if self.get_statusbar_if_error() in (f"Transactie {self.transaction} bestaat niet", f"Transaction {self.transaction} does not exist", f"Transaktion {self.transaction} existiert nicht"):
                self.take_screenshot(screenshot_name="start_transaction_error.jpg")
                raise ValueError(f"Unknown transaction: {self.transaction}")
    
    def end_transaction(self) -> None:
        self.session.endTransaction()
    
    def send_command(self, command: str) -> None:
        try:
            self.session.sendCommand(command)
        except Exception as err:
            self.take_screenshot(screenshot_name="send_command_error.jpg")
            raise ValueError(f"Error sending command {command} -> {err}")

    def click_element(self, id: str = None) -> None:
        if element_type := self.get_element_type(id) in ("GuiTab", "GuiMenu"):
            self.session.findById(id).select()
        elif element_type == "GuiButton":
            self.session.findById(id).press()
        else:
            self.take_screenshot(screenshot_name="click_element_error.jpg")
            raise Warning(f"You cannot use 'Click Element' on element id type {id}")
        self.wait()

    def click_toolbar_button(self, table_id: str, button_id: str) -> None:
        self.element_should_be_present(table_id)
        try:
            self.session.findById(table_id).pressToolbarButton(button_id)
        except AttributeError:
            self.session.findById(table_id).pressButton(button_id)
        except Exception as err:
            self.take_screenshot(screenshot_name="click_toolbar_button_error.jpg")
            raise ValueError(f"Cannot find Table ID/Button ID: {' / '.join([table_id, button_id])}  <-->  {err}")
        self.wait()

    def doubleclick(self, id: str, item_id: str, column_id: str) -> None:
        if element_type := self.get_element_type(id) == "GuiShell":
            self.session.findById(id).doubleClickItem(item_id, column_id)
        else:
            self.take_screenshot(screenshot_name="doubleclick_element_error.jpg")
            raise Warning(f"You cannot use 'doubleclick element' on element type {element_type}")
        self.wait()

    def assert_element_present(self, id: str, message: Optional[str] = None) -> None:
        if not self.is_element(element=id):
            self.take_screenshot(screenshot_name="assert_element_present_error.jpg")
            raise ValueError(message if message is not None else f"Cannot find element {id}")

    def assert_element_value(self, id: str, expected_value: str, message: Optional[str] = None) -> None:
        if self.is_element(element=id):
            actual_value = self.get_value(id=id)
            self.session.findById(id).setfocus()
            self.wait()
        if element_type := self.get_element_type(id) in self.text_elements:
            if expected_value != actual_value:
                message = message if message is not None else f"Element value of {id} should be {expected_value}, but was {actual_value}"
                self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                raise AssertionError(f"Element value of {id} should be {expected_value}, but was {actual_value}")
        elif element_type in ("GuiCheckBox", "GuiRadioButton"):
            if expected_value := bool(expected_value):
                if not actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                    raise AssertionError(f"Element value of {id} should be {expected_value}, but was {actual_value}")
            elif not expected_value:
                if actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                    raise AssertionError(f"Element value of {id} should be {expected_value}, but was {actual_value}")
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
            raise AssertionError(f"Element value of {id} should be {expected_value}, but was {actual_value}")

    def assert_element_value_contains(self, id: str, expected_value: str, message: Optional[str] = None) -> None:
        if self.is_element(element=id):
            actual_value = self.get_value(id=id)
            self.session.findById(id).setfocus()
            self.wait()
        if element_type := self.get_element_type(id) in self.text_elements:
            if expected_value != actual_value:
                message = message if message is not None else f"Element value of {id} does not contain {expected_value} but was {actual_value}"
                self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
                raise AssertionError(message)
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error.jpg")
            raise AssertionError(f"Element value of {id} does not contain {expected_value}, but was {actual_value}")

    def get_cell_value(self, table_id: str, row_num: int, col_id: str) -> str:
        if self.is_element(element=table_id):
            try:
                return self.session.findById(table_id).getCellValue(row_num, col_id)
            except Exception as err:
                self.take_screenshot(screenshot_name="get_cell_value_error.jpg")
                raise ValueError(f"Cannot find cell value for table: {table_id}, row: {row_num}, and column: {col_id} -> {err}")

    def set_combobox(self, id: str, key: str) -> None:
        if element_type := self.get_element_type(id) == "GuiComboBox":
            self.session.findById(id).key = key
            logger.info(f"ComboBox value {key} selected from {id}")
            self.wait()
        else:
            self.take_screenshot(screenshot_name="set_combobox_error.jpg")
            raise ValueError(f"Element type {element_type} for element {id} has no set key method.")

    def get_element_location(self, id: str) -> tuple[int]:
        return (self.session.findById(id).screenLeft, self.session.findById(id).screenTop) if self.is_element(element=id) else (0, 0)

    def get_element_type(self, id) -> object | None:
        try:
            return self.session.findById(id).type if self.is_element(element=id) else None
        except Exception as err:
            self.take_screenshot(screenshot_name="get_element_type_error.jpg")
            raise ValueError(f"Cannot find element type for id: {id} -> {err}")

    def get_row_count(self, table_id) -> int:
        try:
            return self.session.findById(table_id).rowCount if self.is_element(element=table_id) else 0
        except Exception as err:
            self.take_screenshot(screenshot_name="get_row_count_error.jpg")
            raise ValueError(f"Cannot find row count for table: {table_id} -> {err}")

    def get_scroll_position(self, id: str) -> int:
        self.wait()
        try:
            return int(self.session.findById(id).verticalScrollbar.position) if self.is_element(element=id) else 0
        except Exception as err:
            self.take_screenshot(screenshot_name="get_scroll_position_error.jpg")
            raise ValueError(f"Cannot get scrollbar position for: {id} -> {err}")

    def get_window_title(self, id: str) -> str:
        try:
            return self.session.findById(id).text if self.is_element(element=id) else ""
        except Exception as err:
            self.take_screenshot(screenshot_name="get_window_title_error.jpg")
            raise ValueError(f"Cannot find window with locator {id} -> {err}")

    def get_value(self, id: str) -> str | bool:
        try:
            if element_type := self.get_element_type(id) in self.text_elements:
                return self.session.findById(id).text
            elif element_type in ("GuiCheckBox", "GuiRadioButton"):
                return self.session.findById(id).selected
            elif element_type == "GuiComboBox":
                return str(self.session.findById(id).text).strip()
            else:
                self.take_screenshot(screenshot_name="get_value_warning.jpg")
                raise Warning(f"Cannot get value for element type {element_type} for id {id}")
        except Exception as err:
            self.take_screenshot(screenshot_name="get_value_error.jpg")
            raise ValueError(f"Cannot get value for element type {element_type} for id {id} -> {err}")

    def input_text(self, id: str, text: str) -> None:
        if element_type := self.get_element_type(id) in self.text_elements:
            self.session.findById(id).text = text
            if element_type != "GuiPasswordField":
                logger.info(f"Input {text} into text field {id}")
            self.wait()
        else:
            self.take_screenshot(screenshot_name="input_text_error.jpg")
            raise ValueError(f"Cannot use keyword 'input text' for element type {element_type}")

    def string_generator(size: int = 6, chars: str = string.ascii_uppercase + string.digits):
        return ''.join(random.choice(chars) for _ in range(size))

    def input_random_value(self, id: str, text: str, prefix: bool = False, suffix: bool = False, date_time: bool = False, random: bool = False) -> str:
        dt: str = datetime.datetime.now().strftime("%Y%M%d%H%m%s") if date_time else ""
        rs: str = self.string_generator() if random else ""
        tmp: str = text
        if prefix:
            tmp = f"{dt}_{rs}{tmp}"
        if suffix:
            tmp = f"{tmp}_{dt}{rs}"
        self.input_text(id=id, text=tmp)
        return tmp
    
    def input_current_date(self, id: str, format: Optional[str] = "%M/%d/%Y") -> None:
        self.input_text(id=id, text=datetime.datetime.now().strftime(format))

    def maximize_window(self, window: Optional[int] = None) -> None:
        if window:
            self.window = window
        try:
            self.session.findById(f"wnd[{self.window}]").maximize()
            self.wait()
        except Exception as err:
            self.take_screenshot(screenshot_name="maximize_window.jpg")
            raise ValueError(f"Cannot maximize window wnd[{self.window}] -> {err}")

    def set_vertical_scroll(self, id: str, position: int) -> None:
        if self.is_element(id):
            self.session.findById(id).verticalScrollbar.position = position
            self.wait()

    def set_horizontal_scroll(self, id: str, position: int) -> None:
        if self.is_element(id):
            self.session.findById(id).horizontalScrollbar.position = position
            self.wait()

    def get_vertical_scroll(self, id: str) -> int | None:
        return self.session.findById(id).verticalScrollbar.position if self.is_element(id) else None

    def get_horizontal_scroll(self, id: str) -> int | None:
        return self.session.findById(id).horizontalScrollbar.position if self.is_element(id) else None

    def select_checkbox(self, id: str) -> None:
        if element_type := self.get_element_type(id) == "GuiCheckBox":
            self.session.findById(id).selected = True
            self.wait()
        else:
            self.take_screenshot(screenshot_name="select_checkbox_error.jpg")
            raise ValueError(f"Cannot use keyword 'select checkbox' for element type {element_type}")

    def unselect_checkbox(self, id: str) -> None:
        if element_type := self.get_element_type(id) == "GuiCheckBox":
            self.session.findById(id).selected = False
            self.wait()
        else:
            self.take_screenshot(screenshot_name="select_checkbox_error.jpg")
            raise ValueError(f"Cannot use keyword 'unselect checkbox' for element type {element_type}")

    def set_cell_value(self, table_id, row_num, col_id, text):
        if self.is_element(element=table_id):
            try:
                self.session.findById(table_id).modifyCell(row_num, col_id, text)
                logger.info(f"Input {text} into cell ({row_num}, {col_id})")
                self.wait()
            except Exception as err:
                self.take_screenshot(screenshot_name="set_cell_value_error.jpg")
                raise ValueError(f"Failed entering {text} into cell ({row_num}, {col_id}) -> {err}")

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
                    raise ValueError(f"Cannot find given Vkey {vkey}, provide a valid Vkey number or combination")
        try:
            self.session.findById(f"wnd[{self.window}]").sendVKey(vkey_id)
            self.wait()
        except Exception as err:
            self.take_screenshot(screenshot_name="send_vkey_error.jpg")
            raise ValueError(f"Cannot send Vkey to window wnd[{self.window}]]")

    def select_context_menu_item(self, id: str, menu_id: str, item_id: str) -> None:
        if self.is_element(element=id):
            if hasattr(self.session.findById(id), "nodeContextMenu"):
                self.session.findById(id).nodeContextMenu(menu_id)
            elif hasattr(self.session.findById(id), "pressContextButton"):
                self.session.findById(id).pressContextButton(menu_id)
            else:
                self.take_screenshot(screenshot_name="select_context_menu_item_error.jpg")
                raise ValueError(f"Cannot use keyword 'Select Context Menu Item' with element type {self.get_element_type(id)}")
            self.session.findById(id).selectContextMenuItem(item_id)
            self.wait()

    def select_from_list_by_label(self, id: str, value: str) -> None:
        if element_type := self.get_element_type(id) == "GuiComboBox":
            self.session.findById(id).key = value
            self.wait()
        else:
            self.take_screenshot(screenshot_name="select_from_list_by_label_error.jpg")
            raise ValueError(f"Cannot use keyword Select From List By Label with element type {element_type}")

    def select_node(self, tree_id: str, node_id: str, expand: bool = False):
        if self.is_element(element=tree_id):
            self.session.findById(tree_id).selectedNode = node_id
            if expand:
                try:
                    self.session.findById(tree_id).expandNode(node_id)
                except:
                    pass
            self.wait()

    def select_node_link(self, tree_id: str, link_id1: str, link_id2: str) -> None:
        if self.is_element(element=tree_id):
            self.session.findById(tree_id).selectItem(link_id1, link_id2)
            self.session.findById(tree_id).clickLink(link_id1, link_id2)
            self.wait()

    def select_radio_button(self, id: str) -> None:
        if element_type := self.get_element_type(id) == "GuiRadioButton":
            self.session.findById(id).selected = True
        else:
            self.take_screenshot(screenshot_name="select_radio_button_error.jpg")
            raise ValueError(f"Cannot use keyword Select Radio Button with element type {element_type}")
        self.wait()

    def select_table_column(self, table_id: str, column_id: str) -> None:
        if self.is_element(element=table_id):
            try:
                self.session.findById(table_id).selectColumn(column_id)
            except Exception as err:
                self.take_screenshot(screenshot_name="select_table_column_error.jpg")
                raise ValueError(f"Cannot find column ID: {column_id} for table {table_id}")
            self.wait()

    def select_table_row(self, table_id: str, row_num: int):
        if element_type := self.get_element_type(table_id) == "GuiTableControl":
            id = self.session.findById(table_id).getAbsoluteRow(row_num)
            id.selected = -1
        else:
            try:
                self.session.findById(table_id).selectedRows = row_num
            except Exception as err:
                self.take_screenshot(screenshot_name="select_table_row_error.jpg")
                raise ValueError(f"Cannot use keyword Select Table Row for element type {element_type} -> {err}")
        self.wait()