from types import FunctionType
import win32com.client
import re
import string
import random
import base64
import atexit
from typing import Optional, Any
from Utilities.Utilities import *
from SAPLogger.SapLogger import Logger
from SapObject.SapObject import SapObject
from .Decorators import *


class Gui:
    """
     Python Framework library for controlling the SAP GUI Desktop and focused 
     on testing business processes. The library uses the native SAP GUI scripting engine 
     for interaction with the desktop client application.
    """

    __version__: str = '0.0.10'
    __explicit_wait__: float = 0.0

    def __init__(
        self, 
        test_case: Optional[str] = "My Test Case",
        exit_on_error: Optional[bool] = True, 
        screenshot_on_fail: Optional[bool] = True, 
        screenshot_on_pass: Optional[bool] = False, 
        log_path: Optional[str] = None,
        log_file: Optional[str] = None, 
        verbosity: Optional[int] = None, 
        log_format: Optional[str] = None, 
        log_file_mode: Optional[str] = "a", 
        screenshot_dir: Optional[str] = "screenshots", 
        monitor: Optional[int] = 0, 
        explicit_wait: Optional[float] = 0.0, 
        connection_number: Optional[int] = 0, 
        session_number: Optional[int] = 0, 
        connection_name: Optional[str] = None, 
        date_format: Optional[str] = "%m/%d/%Y", 
        close_sap_on_cleanup: Optional[bool] = True, 
        auto_documentation: Optional[bool] = True) -> None:

        atexit.register(self.cleanup)

        self.test_case_name: str = test_case
        self.exit_on_error: bool = exit_on_error
        self.screenshot_on_fail: bool = screenshot_on_fail
        self.screenshot_on_pass: bool = screenshot_on_pass
        self.close_sap_on_cleanup: bool = close_sap_on_cleanup
        self.auto_documentation: bool = auto_documentation
        self.logger: Logger = Logger(log_name=self.test_case_name, log_path=log_path, log_file=log_file, verbosity=verbosity, format=log_format, file_mode=log_file_mode)
        self.__connection_number: int = connection_number
        self.__session_number: int = session_number
        self.connection_name: str = connection_name if connection_name is not None else ""
        self.sap_gui: win32com.client.CDispatch = None
        self.sap_app: win32com.client.CDispatch = None
        self.connection: win32com.client.CDispatch = None
        self.session: win32com.client.CDispatch = None
        self.screenshot_dir: str = screenshot_dir
        self.monitor: int = int(monitor)
        self.screenshot: Screenshot = Screenshot()
        self.date_format: str = str(date_format)

        __explicit_wait__ = explicit_wait

        if not os.path.exists(self.screenshot_dir):
            self.logger.log.debug(f"Screenshot directory {self.screenshot_dir} does not exist, creating it.")
            try:
                os.makedirs(self.screenshot_dir)
            except Exception as err:
                self.logger.log.error(f"Unable to create screenshot directory {self.screenshot_dir} > {err}")
        
        self.screenshot.screenshot_directory = self.screenshot_dir
        self.screenshot.monitor = self.monitor

        self.__window_number: int = 0
        self.window: win32com.client.CDispatch = None
        self.transaction: str = None
        self.sbar: win32com.client.CDispatch = None
        self.session_info: win32com.client.CDispatch = None
        self.children: dict = {}
        self.task_status: str = None
        self.test_status: str = None
        self.test_case_failed: bool = False
        self.failed_tasks: list = []
        self.passed_tasks: list = []
        self.__task: str = ""
        self.element_id: str = None
        self.element: SapObject = None
    
    def transaction_does_not_exist_strings(self) -> tuple:
        return (
            f"Transactie {self.transaction} bestaat niet", 
            f"Transaction {self.transaction} does not exist", 
            f"Transaktion {self.transaction} existiert nicht")
    
    def cleanup(self) -> None:
        if self.test_status is None:
            if self.test_case_failed or self.test_status == FAIL or len(self.failed_tasks) > 0:
                self.test_status = FAIL
            elif not self.test_case_failed and self.test_status != FAIL and len(self.failed_tasks) == 0:
                self.test_status = PASS
            else:
                self.test_status = "UNKNOWN > Check the logs."
        self.documentation(f"{self.test_case_name} completed with status: {self.test_status}")
        if len(self.failed_tasks) > 0:
            self.documentation(str("The following tasks failed: \n" + "\n".join([str(x) for x in self.failed_tasks]) + "\n"))
        if self.close_sap_on_cleanup:
            self.exit(fail_on_error=False, is_task=False)
    
    @property
    def task(self) -> None:
        return self.__task
    
    @task.setter
    def task(self, value: Optional[str] = None) -> None:
        if value:
            self.__task = str(value)
        else:
            self.__task = parent_func()

    def documentation(self, msg: Optional[str] = None) -> None:
        _msg = msg if msg is not None else self.__task
        if _msg is not None and _msg != "":
            self.logger.log.documentation(_msg)

    def is_element(self, element: str) -> bool:
        try:
            self.session.findById(element)
            self.element = self.session.findById(element)
            return True
        except:
            self.logger.log.debug(f"Unable to locate element: {element} ")
            pass
        return False

    @explicit_wait_before(wait_time=__explicit_wait__)
    def take_screenshot(self, screenshot_name: Optional[str] = None, msg: Optional[str] = None) -> None:
        _msg = msg if msg is not None else ""
        _file_names = []
        try:
            try:
                self.window.HardCopy(screenshot_name, "PNG")
            except Exception as err:
                self.logger.log.warning(f"Error while capturing screenshot with SAP GUi HardCopy module, falling back to mss module | {err}")
                if not screenshot_name:
                    _file_names.append(self.screenshot.shot())
                else:
                    _file_names = self.screenshot.shot(name=screenshot_name)
        except Exception as err:
            self.logger.log.error(f"Error while capturing screenshot with mss module | {err}")
        try:
            if _file_names:
                for f in _file_names:
                    encoded_img = None
                    with open(f, "rb") as f_img:
                        encoded_img = base64.b64encode(f_img.read())
                    self.logger.log.shot(f"{_msg}|{f}|{encoded_img}")
        except Exception as err:
            self.logger.log.error(f"Error while encoding screenshot | {err}")

    @explicit_wait_before(wait_time=__explicit_wait__)
    def task_fail(self, msg: Optional[str] = None, ss_name: Optional[str] = None) -> None:
        if msg:
            self.logger.log.error(msg)
        if self.screenshot_on_fail:
            self.take_screenshot(screenshot_name=ss_name, msg=msg)
        self.task_status = FAIL
        self.failed_tasks.append(self.task)
        if self.exit_on_error:
            self.test_status = FAIL
            self.test_case_failed = True
            sys.exit()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def task_pass(self, msg: Optional[str] = None, ss_name: Optional[str] = None) -> None:
        if msg:
            self.logger.log.info(msg)
        if self.screenshot_on_pass:
            self.take_screenshot(screenshot_name=ss_name, msg=msg)
        self.task_status = PASS
        self.passed_tasks.append(self.task)
    
    def pass_fail(self, indicator: Any, template: list[Any], msg_pass: str, msg_fail: str, foe: bool, ss_pass: str, ss_fail: str, invert: Optional[bool] = False):
        if not invert:
            if indicator in template:
                self.task_pass(msg=msg_pass, ss_name=ss_pass)
            else:
                if foe:
                    self.task_fail(msg=msg_fail, ss_name=ss_fail)
                else:
                    self.task_pass(msg=msg_pass, ss_name=ss_pass)
        if invert:
            if indicator not in template:
                self.task_pass(msg=msg_pass, ss_name=ss_pass)
            else:
                if foe:
                    self.task_fail(msg=msg_fail, ss_name=ss_fail)
                else:
                    self.task_pass(msg=msg_pass, ss_name=ss_pass)
    
    def on_error(self, error: str, msg: str, task: bool, foe: bool, ss: str) -> None:
        _msg = f"{msg}|{error}"
        if task:
            if foe:
                self.task_fail(msg=_msg, ss_name=ss)
            else:
                self.task_pass(msg=_msg, ss_name=ss)
        else:
            self.logger.log.warning(_msg)

    def parse_children(self, parent: Optional[win32com.client.CDispatch] = None, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if parent and type(parent) is win32com.client.CDispatch:
            try:
                for i in parent.children:
                    self.children[i.id] = {
                        "id": i.id, 
                        "text": i.text, 
                        "type": i.type, 
                        "name": i.name, 
                        "screen_left": i.screenLeft, 
                        "screen_top": i.screenTop, 
                        "left": i.left, 
                        "top": i.top, 
                        "tooltip": i.tooltip, 
                        "height": i.height, 
                        "width": i.width
                        }
                    try:
                        self.parse_children(parent=i)
                    except:
                        pass
                return None
            except Exception as err:
                if is_task:
                    if fail_on_error: 
                        self.task_fail(msg=f"Unknown parent: {parent.id} -> {err}")
                    else:
                        self.task_pass(msg=f"Unknown parent: {parent.id}|{err}")
        elif self.element and type(self.element) is win32com.client.CDispatch:
            try:
                for i in self.element.children:
                    self.children[i.id] = {
                        "id": i.id, 
                        "text": i.text, 
                        "type": i.type, 
                        "name": i.name, 
                        "screen_left": i.screenLeft, 
                        "screen_top": i.screenTop, 
                        "left": i.left, 
                        "top": i.top, 
                        "tooltip": i.tooltip, 
                        "height": i.height, 
                        "width": i.width
                        }
                    try:
                        self.parse_children(parent=i)
                    except:
                        pass
                return None
            except Exception as err:
                if is_task:
                    if fail_on_error: 
                        self.task_fail(msg=f"Unknown element: {self.element.id} -> {err}")
                    else:
                        self.task_pass(msg=f"Unknown element: {self.element.id}|{err}")
        else:
            if is_task:
                if fail_on_error: 
                    self.task_fail(msg=f"Not a parent or element or parent/element is not a valid win32com.client.CDispatch object -> {err}")
                else:
                    self.task_pass(msg=f"Not a parent or element or parent/element is not a valid win32com.client.CDispatch object|{err}")
            return None

    def get_element_type(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> str | None:
        _tmp = None
        if is_task: self.task
        try:
            _tmp = self.session.findById(id).type
            if is_task:
                self.task_pass()
        except Exception as err:
            if is_task:
                if fail_on_error: 
                    self.task_fail(msg=f"Unknown element id: {id} -> {err}")
                else:
                    self.task_pass(msg=f"Unknown element id: {id}|{err}")
        return _tmp

    def connect_to_session(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        """
        Connects to a SAP session.

        Keyword Arguments:
            fail_on_error {Optional[bool]} -- Fail test case if an error occurs. (default: {True})
            is_task {Optional[bool]} -- If the current execution if a user defined task for a test case or not. (default: {True})

        Raises:
            ConnectionError: Error while getting SAP GUI object using win32com.client
            ConnectionError: Error while getting SAP scripting engine
            ConnectionError: Error while getting SAP connection to Window 
            ConnectionError: SAP scripting is disable for this server
            ConnectionError: Error while getting SAP session to Window
            ConnectionError: Unable to get status bar during session connection
            ConnectionError: Unable to get session information
        """
        if is_task:
            self.task
        try:
            self.sap_gui = win32com.client.GetObject("SAPGUI")
            if not type(self.sap_gui) == win32com.client.CDispatch:
                raise ConnectionError("Error while getting SAP GUI object using win32com.client")
            self.sap_app = self.sap_gui.GetScriptingEngine
            if not type(self.sap_app) == win32com.client.CDispatch:
                self.sap_gui = None
                raise ConnectionError("Error while getting SAP scripting engine")
            self.connection = self.sap_app.Children(self.__connection_number)
            if not type(self.connection) == win32com.client.CDispatch:
                self.sap_app = None
                self.sap_gui = None
                raise ConnectionError(f"Error while getting SAP connection to Window {self.__connection_number}")
            if self.connection.DisabledByServer == True:
                self.logger.log.error("SAP scripting is disable for this server")
                self.sap_app = None
                self.sap_gui = None
                raise ConnectionError("SAP scripting is disable for this server")
            self.session = self.connection.Children(self.__session_number)
            if not type(self.session) == win32com.client.CDispatch:
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                raise ConnectionError(f"Error while getting SAP session to Window {self.__session_number}")
            if self.session.Info.IsLowSpeedConnection == True:
                self.logger.log.error("SAP connection is listed as low speed, scripting not possible")
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                raise ConnectionError("SAP connection is listed as low speed, scripting not possible")
            self.window = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.__window_number}]")
            if not type(self.window) == win32com.client.CDispatch:
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                self.session = None
                raise ConnectionError("Unable to get window during session connection")
            self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.__window_number}]/sbar")
            if not type(self.sbar) == win32com.client.CDispatch:
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                self.session = None
                raise ConnectionError("Unable to get status bar during session connection")
            self.session_info = self.session.info
            if self.session_info is None:
                self.connection = None
                self.sap_app = None
                self.sap_gui = None
                self.session = None
                self.session_info = None
                raise ConnectionError("Unable to get session information")
            if is_task:
                self.task_pass()
        except Exception as err:
            _msg = f"Unknown error while establishing connection with SAP GUI|{sys.exc_info()[0]}|{err}"
            if is_task:
                if fail_on_error:
                    self.task_fail(_msg)
                else:
                    self.task_pass(_msg)
            else:
                self.logger.log.warning(_msg)

    def connect_to_existing_connection(self, connection_name: Optional[str] = None, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        """
        Connect to an existing SAP connection and get a session.

        Keyword Arguments:
            connection_name {Optional[str]} -- The name of the SAP connect from SAP Logon Pad (default: {None})
            fail_on_error {Optional[bool]} -- If case should fail if there is an error during the execution (default: {True})
            is_task {Optional[bool]} -- If the current function call is a task called by the user (default: {True})
        """
        if is_task:
            self.task
        if connection_name:
            self.connection_name = connection_name
        try:
            self.connection = self.sap_gui.Children(self.__connection_number)
            if self.connection.Description == self.connection_name:
                self.session = self.connection.children(self.session_number)
                self.wait(2.0)
                self.window = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.__window_number}]")
                self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.session_number}]/wnd[{self.__window_number}]/sbar")
                self.session_info = self.session.info
                self.task_pass()
            else:
                self.task_fail(msg=f"No existing connection for {self.connection_name} found.", ss_name="connect_to_existing_connection_error")
        except Exception as err:
            _msg = f"Unknown error while trying to establish existing connection for {self.connection_name}|{err}."
            if is_task:
                if fail_on_error:
                    self.task_fail(msg=_msg, ss_name="connect_to_existing_connection_error")
                else:
                    self.task_pass(msg=_msg, ss_name="connect_to_existing_connection_error")
            else:
                self.logger.log.warning(_msg)

    def open_connection(self, connection_name: Optional[str] = None, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        """
        Open a new SAP connection and session.

        Keyword Arguments:
            connection_name {Optional[str]} -- The name of the SAP connection from SAP Logon Pad (default: {None})
            fail_on_error {Optional[bool]} -- If case should fail if there is an error during the execution (default: {True})
            is_task {Optional[bool]} -- If the current function call is a task called by the user (default: {True})
        """
        self.connection_name = connection_name if connection_name else self.connection_name
        if self.auto_documentation:
            self.documentation(msg=f"Opening connection for {self.connection_name}")
        if is_task:
            self.task
        if not hasattr(self.sap_app, "OpenConnection"):
            try:
                self.sap_gui = win32com.client.GetObject("SAPGUI")
                if not type(self.sap_gui) == win32com.client.CDispatch:
                    self.task_fail("Error while getting SAP GUI object using win32com.client")
                self.sap_app = self.sap_gui.GetScriptingEngine
                if not type(self.sap_app) == win32com.client.CDispatch:
                    self.sap_gui = None
                    self.task_fail("Error while getting SAP scripting engine")
                self.connection = self.sap_app.OpenConnection(self.connection_name, True)
                self.session = self.connection.children(self.__session_number)
                self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.__window_number}]/sbar")
                self.session_info = self.session.info
                self.task_pass(ss_name="open_connection")
                if self.auto_documentation:
                    self.documentation(msg=f"Connection open for {self.connection_name}")
            except Exception as err:
                _msg = f"Cannot open connection {self.connection_name}, please check connection name|{err}"
                if is_task:
                    if fail_on_error:
                        self.task_fail(msg=_msg, ss_name="open_connection")
                    else:
                        self.task_pass(msg=_msg, ss_name="open_connection")
                else:
                    self.logger.log.warning(_msg)

    def get_status_msg_dict(self) -> dict:
        """
        Gets the SAP status message text as a dictionary.

        Returns:
            dict -- dict of status message
        """
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

    def exit(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        """
        Exit the current SAP session.

        Keyword Arguments:
            fail_on_error {Optional[bool]} -- If case should fail if there is an error during the execution (default: {True})
            is_task {Optional[bool]} -- If the current function call is a task called by the user (default: {True})
        """
        if is_task:
            self.task
        try:
            self.connection.closeSession(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]")
            self.connection.closeConnection()
            self.task_pass(msg="Exit successfully.", ss_name="exit")
        except Exception as err:
            _msg = f"Unknown error while attempting to exit SAP session.|{err}"
            if is_task:
                if fail_on_error:
                    self.task_fail(msg=_msg, ss_name="exit")
                else:
                    self.task_pass(msg=_msg, ss_name="exit")
            else:
                self.logger.log.warning(_msg)

    def maximize_window(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        """
        Maximize the current SAP window to fullsize.

        Keyword Arguments:
            fail_on_error {Optional[bool]} -- If case should fail if there is an error during the execution (default: {True})
            is_task {Optional[bool]} -- If the current function call is a task called by the user (default: {True})
        """
        if is_task:
            self.task
        try:
            # self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.__window_number}]").maximize()
            self.window.maximize()
            self.task_pass(msg="Maximize of SAP window successful", ss_name="maximize_window")
        except Exception as err:
            _msg = f"Unknown error while attempting to maximize SAP window.|{err}"
            if is_task:
                if fail_on_error:
                    self.task_fail(msg=_msg, ss_name="maximize_window")
                else:
                    self.task_pass(msg=_msg, ss_name="maximize_window")
            else:
                self.logger.log.warning(_msg)

    def restart_session(self, connection_name: Optional[str] = None, delay: Optional[float] = 0.0, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        """
        Restart the SAP session by exiting and reopening a new session.

        Arguments:
            connection_name {str} -- The name of the SAP connect from SAP Logon Pad

        Keyword Arguments:
            delay {Optional[float]} -- Additional delay after reopening SAP session so it can correctly load before proceeding. (default: {0.0})
            fail_on_error {Optional[bool]} -- If case should fail if there is an error during the execution (default: {True})
            is_task {Optional[bool]} -- If the current function call is a task called by the user (default: {True})
        """
        if is_task:
            self.task
        if connection_name is not None:
            self.connection_name = connection_name
        try:
            self.exit()
            self.open_connection(connection_name=self.connection_name, fail_on_error=False, is_task=False)
            self.wait(value=delay)
            self.maximize_window()
            self.task_pass(msg="Successfully restart SAP session.", ss_name="restart_session")
        except Exception as err:
            _msg = f"Unknown error while attempting to restart SAP session.|{err}"
            if is_task:
                if fail_on_error:
                    self.task_fail(msg=_msg, ss_name="restart_session")
                else:
                    self.task_pass(msg=_msg, ss_name="restart_session")
            else:
                self.logger.log.warning(_msg)

    def wait_for_element(self, id: str, timeout: Optional[float] = 60.0, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        """
        Checks every 0.5 seconds for a SAP GUI element to become available and returning once the element is available or the timeout expairs.  

        Arguments:
            id {str} -- SAP GUI element's if string

        Keyword Arguments:
            timeout {Optional[float]} -- Amount of time to wait for element to become available (default: {60.0})
            fail_on_error {Optional[bool]} -- If case should fail if there is an error during the execution (default: {True})
            is_task {Optional[bool]} -- If the current function call is a task called by the user (default: {True})
        """
        if is_task: self.task
        t = Timer()
        while True:
            if not self.is_element(element=id) and t.elapsed() <= timeout:
                self.wait(value=0.5)
            else:
                break
        if not self.is_element(element=id):
            if is_task:
                if fail_on_error:
                    self.task_fail(msg=f"Wait For Element could not find element with id {id}", ss_name="wait_for_element_error")
                else:
                    self.task_pass(msg=f"Wait For Element could not find element with id {id}", ss_name="wait_for_element")
            else:
                self.logger.log.warning(f"Wait For Element could not find element with id {id}")
        else:
            self.task_pass(msg=f"Wait For Element with id {id} successful", ss_name="wait_for_element")

    def get_statusbar_if_error(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> str | None:
        """
        Get SAP statusbar message if statusbar in error state.

        Returns:
            str | None -- Statusbar error text or None
        """
        _tmp = None
        if is_task:
            self.task
        try:
            if self.sbar.messageType == "E":
                _tmp = f"{self.sbar.findById('pane[0]').text} -> Message no. {self.sbar.messageId.strip('')}:{self.sbar.messageNumber}"
            self.task_pass(msg=f"get_statusbar_if_error was successful", ss_name="get_statusbar_if_error")
        except Exception as err:
            _msg = f"Unhandled error while checking if statusbar had error msg.|{err}"
            if is_task:
                if fail_on_error:
                    self.task_fail(msg=_msg, ss_name="error_for_get_statusbar_if_error")
                else:
                    self.task_pass(msg=_msg, ss_name="error_for_get_statusbar_if_error")
            else:
                self.log.warning(_msg)
        return _tmp

    @explicit_wait_after(wait_time=__explicit_wait__)
    def start_transaction(self, transaction: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.transaction = transaction.upper()
            self.session.startTransaction(self.transaction)
            s_msg = str(self.sbar.findById('pane[0]').text).strip(" \n\r\t")
            _template = [i for i in self.transaction_does_not_exist_strings()]
            self.pass_fail(
                indicator=s_msg, 
                template=_template, 
                msg_pass=f"Started transaction {self.transaction} successfully|{s_msg}", 
                msg_fail=f"ValueError|{s_msg}", 
                foe=fail_on_error, 
                ss_pass="start_transaction", 
                ss_fail="start_transaction_error", 
                invert=True)
        except Exception as err:
            self.on_error(err, f"Unhandled error while starting transaction {transaction}|{err}", is_task, fail_on_error, ss="start_transaction_error")
    
    start: FunctionType = start_transaction
    Start: FunctionType = start_transaction
    START: FunctionType = start_transaction
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def end_transaction(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.session.endTransaction()
            self.task_pass()
        except Exception as err:
            self.on_error(err, "Error ending transaction", is_task, fail_on_error, ss="end_transaction_error")
    
    end: FunctionType = end_transaction
    End: FunctionType = end_transaction
    END: FunctionType = end_transaction
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def send_command(self, command: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.session.sendCommand(command)
            self.task_pass()
        except Exception as err:
            self.on_error(err, f"Error sending command {command}", is_task, fail_on_error, ss="send_command_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def click_element(self, id: str = None, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if (element_type := self.get_element_type(id)) in ("GuiTab", "GuiMenu"):
                self.session.findById(id).select()
                self.task_pass(msg=f"Clicking element: {id} was successful", ss_name="click_element_success")
            elif element_type == "GuiButton":
                self.session.findById(id).press()
                self.task_pass(msg=f"Clicking GuiButton with: {id} was successful", ss_name="click_gui_button_success")
            else:
                self.task_fail(msg=f"Clicking element: {id} failed", ss_name="click_element_failed")
        except Exception as err:
            self.on_error(err, f"Unknown error while clicking element: {id}", is_task, fail_on_error, ss="click_element_error")
    
    click: FunctionType = click_element

    @explicit_wait_after(wait_time=__explicit_wait__)
    def click_toolbar_button(self, table_id: str, button_id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.element_should_be_present(table_id)
            self.session.findById(table_id).pressToolbarButton(button_id)
            self.task_pass(msg=f"Clicking toolbar button: {id} was successful", ss_name="click_toolbar_button_success")
        except AttributeError:
            try:
                self.session.findById(table_id).pressButton(button_id)
                self.task_pass(msg=f"Clicking toolbar button: {id} was successful", ss_name="click_toolbar_button_success")
            except Exception as err:
                self.on_error(err, f"Cannot find Table ID/Button ID: {' / '.join([table_id, button_id])}", is_task, fail_on_error, ss="click_toolbar_button_error")
        except Exception as err:
            self.on_error(err, f"Cannot find Table ID/Button ID: {' / '.join([table_id, button_id])}", is_task, fail_on_error, ss="click_toolbar_button_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def double_click(self, id: str, item_id: str, column_id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        element_type = None
        try:
            element_type = self.get_element_type(id)
        except Exception as err:
            self.on_error(err, f"Error while getting element type for ID: {id}", is_task, fail_on_error, ss="double_click_get_type_error")
        try:
            if element_type == "GuiShell":
                self.session.findById(id).doubleClickItem(item_id, column_id)
                self.task_pass(msg=f"Double clicking: {id} was successful", ss_name="double_click_success")
            else:
                try:
                    self.session.findById(id).doubleClickItem(item_id, column_id)
                    self.task_pass(msg=f"Double clicking: {id} of type {element_type} was successful", ss_name="double_click_success")
                except Exception as err:
                    self.on_error(err, f"You cannot use 'double_click' for element: {id} of type {element_type}", is_task, fail_on_error, ss="double_click_element_error")
        except Exception as err:
            self.on_error(err, f"You cannot use 'double_click' for element: {id} of type {element_type}", is_task, fail_on_error, ss="double_click_element_error")
        
    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_cell_value(self, table_id: str, row_num: int, col_id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> str | None:
        if is_task: self.task
        if self.is_element(element=table_id):
            try:
                _value = self.session.findById(table_id).getCellValue(row_num, col_id)
                self.task_pass()
                return _value
            except Exception as err:
                self.on_error(err, f"Cannot find cell value for table: {table_id}, row: {row_num}, and column: {col_id}", is_task, fail_on_error, ss="get_cell_value_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def set_combobox(self, id: str, key: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if (element_type := self.get_element_type(id)) == "GuiComboBox":
                self.session.findById(id).key = key
                self.logger.log.info(f"ComboBox value {key} selected from {id}")
                self.task_pass()
            else:
                self.on_error(None, f"Element type {element_type} for element {id} has no set key method.", is_task, fail_on_error, ss="set_combobox_error")
        except Exception as err:
            self.on_error(err, f"Unknown error while setting ComboBox value {key} for {id}", is_task, fail_on_error, ss="set_combobox_error")
    
    combobox = set_combobox
    set_dropdown = set_combobox

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_element_location(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> tuple[int] | None:
        if is_task: self.task
        _location = None
        try:
            _location = (self.session.findById(id).screenLeft, self.session.findById(id).screenTop) if self.is_element(element=id) else None
            if _location:
                self.task_pass()
            else:
                self.task_fail()
        except Exception as err:
            self.on_error(err, f"Unknown error while getting element location for {id}", is_task, fail_on_error, ss="get_element_location_error")
        return _location

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_element_type(self, id) -> Any:
        try:
            _type = self.session.findById(id).type
            self.task_pass()
            return _type
        except Exception as err:
            self.take_screenshot(screenshot_name="get_element_type_error")
            self.logger.log.error(f"Cannot find element type for id: {id} -> {err}")
            self.task_fail()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_row_count(self, table_id) -> int:
        try:
            _count = self.session.findById(table_id).rowCount if self.is_element(element=table_id) else 0
            self.task_pass()
            return _count
        except Exception as err:
            self.take_screenshot(screenshot_name="get_row_count_error")
            self.logger.log.error(f"Cannot find row count for table: {table_id} -> {err}")
            self.task_fail()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_scroll_position(self, id: str) -> int:
        try:
            _position = int(self.session.findById(id).verticalScrollbar.position) if self.is_element(element=id) else 0
            self.task_pass()
            return _position
        except Exception as err:
            self.take_screenshot(screenshot_name="get_scroll_position_error")
            self.logger.log.error(f"Cannot get scrollbar position for: {id} -> {err}")
            self.task_fail()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_window_title(self, id: str) -> str:
        try:
            _title =  self.session.findById(id).text if self.is_element(element=id) else ""
            self.task_pass()
            return _title
        except Exception as err:
            self.take_screenshot(screenshot_name="get_window_title_error")
            self.logger.log.error(f"Cannot find window with locator {id} -> {err}")
            self.task_fail()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_value(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> Any:
        if is_task: self.task
        try:
            _value = None
            if (element_type := self.get_element_type(id)) in text_elements:
                _value = self.session.findById(id).text
            elif element_type in ("GuiCheckBox", "GuiRadioButton"):
                _value = self.session.findById(id).selected
            elif element_type == "GuiComboBox":
                _value = str(self.session.findById(id).text).strip()
            else:
                self.take_screenshot(screenshot_name="get_value_warning")
                self.logger.log.error(f"Cannot get value for element type {element_type} for id {id}")
            if _value:
                self.task_pass()
                return _value
            else:
                return None
        except Exception as err:
            _msg = f"Cannot get value for element type {element_type} for id {id}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="get_value_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def input_text(self, id: str, text: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if (element_type := self.get_element_type(id)) in text_elements:
                self.session.findById(id).text = text
                self.task_pass(msg=f"Input {text} into text field {id} was successful", ss_name="input_text_passed")
            else:
                self.task_fail(msg=f"Cannot use keyword 'input text' for element type {element_type}", ss_name="input_text_failed")
        except Exception as err:
            _msg = f"Unknown error during text input for id: {id}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="input_text_error")
    
    text = input_text

    def input_random_value(self, id: str, text: Optional[str] = None, prefix: Optional[bool] = False, suffix: Optional[bool] = False, 
        date_time: Optional[bool] = False, random: Optional[bool] = False, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> str:
        if is_task: self.task
        dt: str = str("_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")) if date_time else ""
        rs: str = str("_" + string_generator()) if random else ""
        tmp: str = str("_" + text) if text is not None else ""
        if prefix:
            tmp = f"{dt}{rs}{tmp}"
        if suffix:
            tmp = f"{tmp}{dt}{rs}"
        try:
            self.input_text(id=id, text=tmp, is_task=False)
            if self.get_value(id=id) == tmp:
                self.task_pass(msg=f"Input random value: {tmp} was successful", ss_name="input_random_value_passed")
            else:
                self.task_fail(msg=f"Input random value: {tmp} failed", ss_name="input_random_value_failed")
        except Exception as err:
            _msg = f"Unknown error inputting random value: {tmp}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="input_random_value_error")
        return tmp
    
    def input_current_date(self, id: str, format: Optional[str] = "%m/%d/%Y", fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        self.input_text(id=id, text=datetime.datetime.now().strftime(format), is_task=False)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def set_vertical_scroll(self, id: str, position: int, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        if self.is_element(id):
            self.session.findById(id).verticalScrollbar.position = position
            self.task_pass()
        else:
            self.task_fail()

    @explicit_wait_after(wait_time=__explicit_wait__)
    def set_horizontal_scroll(self, id: str, position: int, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        if self.is_element(id):
            self.session.findById(id).horizontalScrollbar.position = position
            self.task_pass()
        else:
            self.task_fail()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_vertical_scroll(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> int | None:
        if is_task: self.task
        try:
            _vs = self.session.findById(id).verticalScrollbar.position if self.is_element(id) else None
            self.task_pass()
            return _vs
        except Exception as err:
            self.take_screenshot(screenshot_name="get_vertical_scroll")
            self.logger.log.error(f"Cannot get vertical scroll position -> {err}")
            self.task_fail()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_horizontal_scroll(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> int | None:
        if is_task: self.task
        try:
            _hs = self.session.findById(id).horizontalScrollbar.position if self.is_element(id) else None
            self.task_pass()
            return _hs
        except Exception as err:
            self.take_screenshot(screenshot_name="get_horizontal_scroll")
            self.logger.log.error(f"Cannot get horizontal scroll position -> {err}")
            self.task_fail()

    @explicit_wait_after(wait_time=__explicit_wait__)
    def select_checkbox(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        if (element_type := self.get_element_type(id)) == "GuiCheckBox":
            self.session.findById(id).selected = True
            self.task_pass()
        else:
            self.take_screenshot(screenshot_name="select_checkbox_error")
            self.logger.log.error(f"Cannot use keyword 'select checkbox' for element type {element_type}")
            self.task_fail()

    @explicit_wait_after(wait_time=__explicit_wait__)
    def unselect_checkbox(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        if (element_type := self.get_element_type(id)) == "GuiCheckBox":
            self.session.findById(id).selected = False
            self.task_pass()
        else:
            self.take_screenshot(screenshot_name="select_checkbox_error")
            self.logger.log.error(f"Cannot use keyword 'unselect checkbox' for element type {element_type}")
            self.task_fail()

    @explicit_wait_after(wait_time=__explicit_wait__)
    def set_cell_value(self, table_id, row_num, col_id, text, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True):
        if is_task: self.task
        try:
            if self.is_element(element=table_id):
                self.session.findById(table_id).modifyCell(row_num, col_id, text)
                self.task_pass(msg=f"Input {text} into cell ({row_num}, {col_id}) was successful", ss_name="set_cell_value_passed")
            else:
                self.task_fail(msg=f"Failed entering {text} into cell ({row_num}, {col_id})", ss_name="set_cell_value_failed")
        except Exception as err:
            _msg = f"Failed entering {text} into cell ({row_num}, {col_id})|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="set_cell_value_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def send_vkey(self, vkey: str, window: Optional[int] = 0, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        vkey_id = str(vkey)
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
                    self.task_fail(msg=f"Cannot find given Vkey {vkey}, provide a valid Vkey number or combination", ss_name="find_vkey_failed")
        try:
            self.window.sendVKey(vkey_id)
            self.task_pass(msg=f"Send {vkey} successful", ss_name="send_vkey_passed")
        except Exception as err:
            _msg = f"Cannot send Vkey to window wnd[{self.__window_number}]|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="send_vkey_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def select_context_menu_item(self, id: str, menu_id: str, item_id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if self.is_element(element=id):
                if hasattr(self.session.findById(id), "nodeContextMenu"):
                    self.session.findById(id).nodeContextMenu(menu_id)
                    self.session.findById(id).selectContextMenuItem(item_id)
                    self.task_pass(msg=f"Selecting item: {item_id} from menu: {menu_id} for ID: {id} was successful", ss_name="select_context_menu_item_passed")
                elif hasattr(self.session.findById(id), "pressContextButton"):
                    self.session.findById(id).pressContextButton(menu_id)
                    self.session.findById(id).selectContextMenuItem(item_id)
                    self.task_pass(msg=f"Selecting item: {item_id} from menu: {menu_id} for ID: {id} was successful", ss_name="select_context_menu_item_passed")
                else:
                    self.task_fail(msg=f"Cannot use keyword 'Select Context Menu Item' with element type {self.get_element_type(id)}", ss_name="select_context_menu_item_failed")
        except Exception as err:
            _msg = f"Unknown error while selecting item: {item_id} from menu: {menu_id} for ID: {id}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="select_context_menu_item_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def select_from_list_by_label(self, id: str, value: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if (element_type := self.get_element_type(id)) == "GuiComboBox":
                self.session.findById(id).key = value
                self.task_pass(msg=f"Selecting item from list: {id} by label: {value} was successful", ss_name="select_from_list_by_label_passed")
            else:
                self.task_fail(msg=f"Cannot use keyword Select From List By Label with element type {element_type}", ss_name="select_from_list_by_label_failed")
        except Exception as err:
            _msg = f"Unknown error while selecting item from list: {id} by label: {value}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="select_from_list_by_label_error")

    @explicit_wait_before(wait_time=__explicit_wait__)
    def select_node(self, tree_id: str, node_id: str, expand: bool = False, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True):
        if is_task: self.task
        try:
            if self.is_element(element=tree_id):
                self.session.findById(tree_id).selectedNode = node_id
                if expand:
                    try:
                        self.session.findById(tree_id).expandNode(node_id)
                        self.task_pass(msg=f"Selecting node {node_id} from tree {tree_id} was successful", ss_name="select_node_passed")
                    except:
                        _msg = f"Unknown error while selecting node {node_id} from tree {tree_id}|{err}"
                        self.on_error(err, _msg, is_task, fail_on_error, ss="select_node_error")
            else:
                self.task_fail(msg=f"Unable to select node {node_id} from tree {tree_id}", ss_name="select_node_failed")
        except Exception as err:
            _msg = f"Unknown error while selecting node {node_id} from tree {tree_id}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="select_node_error")

    @explicit_wait_before(wait_time=__explicit_wait__)
    def select_node_link(self, tree_id: str, link_id1: str, link_id2: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if self.is_element(element=tree_id):
                self.session.findById(tree_id).selectItem(link_id1, link_id2)
                self.session.findById(tree_id).clickLink(link_id1, link_id2)
                self.task_pass(msg=f"Selecting node {link_id1} and clicked link {link_id2} from tree {tree_id}", ss_name="select_node_link_passed")
            else:
                self.task_fail(msg=f"Unable to select node {link_id1} and click link {link_id2} from tree {tree_id}", ss_name="select_node_link_failed")
        except Exception as err:
            _msg = f"Unknown error while selecting node {link_id1} and clicking link {link_id2} from tree {tree_id}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="select_node_link_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def select_radio_button(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if (element_type := self.get_element_type(id)) == "GuiRadioButton":
                self.session.findById(id).selected = True
                self.task_pass(msg=f"Select Radio Button: {id} was successful", ss_name="select_radio_button_passed")
            else:
                self.task_fail(msg=f"Cannot use keyword Select Radio Button with element type {element_type}", ss_name="select_radio_button_failed")
        except Exception as err:
            _msg = f"Unknown error while selecting Radio Button: {id}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="select_radio_button_error")

    # Function for Tables
    @explicit_wait_after(wait_time=__explicit_wait__)
    def select_table_row(self, table_id: str, row_num: int, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if (element_type := self.get_element_type(table_id)) == "GuiTableControl":
                id = self.session.findById(table_id).getAbsoluteRow(row_num)
                id.selected = -1
                self.task_pass(msg=f"Selecting row: {row_num} from table: {table_id} was successful", ss_name="select_table_row_passed")
            else:
                try:
                    self.session.findById(table_id).selectedRows = row_num
                    self.task_pass(msg=f"Selecting row: {row_num} from table: {table_id} was successful", ss_name="select_table_row_passed")
                except Exception as err:
                    _msg = f"Cannot use keyword Select Table Row for element type {element_type}|{err}"
                    self.on_error(err, _msg, is_task, fail_on_error, ss="select_table_row_error")
        except Exception as err:
            _msg = f"Cannot use keyword Select Table Row for element type {element_type}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="select_table_row_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def get_next_empty_table_row(self, table_id: str, column_index: Optional[int] = 0, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            table = self.session.findById(table_id)
            rows = table.rows
            for i in range(rows.count):
                row = rows.elementAt(i)
                if row.elementAt(column_index).text == "":
                    self.task_pass(msg=f"Found next empty table row for table {table_id} at row: {i}", ss_name="get_next_empty_table_row_passed")
                    return i
            self.task_fail(msg=f"Cannot get next empty table row for table {table_id}", ss_name="get_next_empty_table_row_failed")
        except Exception as err:
            _msg = f"Unknown error while getting next empty table row for table {table_id}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="get_next_empty_table_row_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def insert_in_table(self, table_id: str, value: str, column_index: Optional[int] = 0, row_index: Optional[int] = None, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if not row_index:
                row_index = self.get_next_empty_table_row(table_id=table_id, column_index=column_index, is_task=False)
            table = self.session.findById(table_id)
            cell = table.getCell(row_index, column_index)
            (element_type := cell.type)
            if (element_type := cell.type) == "GuiComboBox":
                cell.key = value
                self.task_pass(msg=f"Inserting value: {value} in table: {table_id} was successful", ss_name="insert_in_table_passed")
            elif element_type == "GuiCTextField":
                cell.text = value
                self.task_pass(msg=f"Inserting value: {value} in table: {table_id} was successful", ss_name="insert_in_table_passed")
            else:
                self.task_fail(msg=f"Cannot inset {value} in table {table_id} at column index: {column_index}", ss_name="insert_in_table_failed")
        except Exception as err:
            _msg = f"Unknown error while inserting value: {value} in table: {table_id} at column index: {column_index}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="insert_in_table_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def select_table_column(self, table_id: str, column_id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if self.is_element(element=table_id):
                self.session.findById(table_id).selectColumn(column_id)
                self.task_pass(msg=f"Selecting column: {column_id} from table: {table_id} was successful", ss_name="select_table_column_passed")
            else:
                self.task_fail(msg=f"Selecting column: {column_id} from table: {table_id} failed", ss_name="get_next_empty_table_row_failed")
        except Exception as err:
            _msg = f"Cannot find column ID: {column_id} for table {table_id}|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="select_table_column_error")

    def dump_grid_view(self, table_id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> list:
        if is_task: self.task
        _rows = []
        try:
            _table = self.session.findById(table_id)
            _row_count = _table.rowCount
            _column_names = [i for i in _table.columnOrder]
            for row in range(_row_count): 
                _cells = []
                for column in _column_names:
                    _cells.append(_table.getCellValue(row, column))
                _rows.append(_cells)
        except Exception as err:
            _msg = f"Unknown error in dump_grid_view|{err}"
            self.on_error(err, _msg, is_task, fail_on_error, ss="dump_grid_view_error")
        if len(_rows) != 0:
            self.task_pass(msg=f"GridView dump successfully", ss_name="dump_grid_view_pass_passed")
        else:
            self.task_fail(msg=f"GridView dump failed", ss_name="dump_grid_view_pass_failed")
        return _rows

    # Soft functions
    def try_and_continue(self, func_name: str, *args, **kwargs) -> Any:
        result = None
        try:
            if hasattr(self, func_name) and callable(func := getattr(self, func_name)):
                result = func(*args, **kwargs)
        except Exception as err:
            self.logger.log.debug(f"Try and Continue error for function: {func}|{err}")
        return result
    
    # Buttons & Keys
    @explicit_wait_after(wait_time=__explicit_wait__)
    def enter(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="ENTER", is_task=False)
        except Exception as err:
            _msg = "Error while pressing ENTER"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_enter_error")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def save(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="CTRL+S")
        except Exception as err:
            _msg = "Error during SAVE"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_save_error")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def back(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F3")
        except Exception as err:
            _msg = "Error while pressing BACK"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_back_error")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def f8(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F8")
        except Exception as err:
            _msg = "Error while pressing F8"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_f8_error")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def f5(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F5")
        except Exception as err:
            _msg = "Error while pressing F5"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_f5_error")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def f6(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F6")
        except Exception as err:
            _msg = "Error while pressing F6"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_f6_error")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def f7(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F7")
        except Exception as err:
            _msg = "Error while pressing F7"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_f7_error")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def f4(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F4")
        except Exception as err:
            _msg = "Error while pressing F4"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_f4_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f3(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F3")
        except Exception as err:
            _msg = "Error while pressing F3"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_f3_error")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def f2(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F2")
        except Exception as err:
            _msg = "Error while pressing F2"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_f2_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f1(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            self.send_vkey(vkey="F1")
        except Exception as err:
            _msg = "Error while pressing F1"
            self.on_error(err, _msg, is_task, fail_on_error, ss="press_f1_error")

    # Assertions
    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_element_value_contains(self, id: str, expected_value: str, message: Optional[str] = None) -> None:
        if self.is_element(element=id):
            actual_value = self.get_value(id=id)
            self.session.findById(id).setfocus()
            self.task_pass()
        if (element_type := self.get_element_type(id)) in text_elements:
            if expected_value != actual_value:
                message = message if message is not None else f"Element value of {id} does not contain {expected_value} but was {actual_value}"
                self.take_screenshot(screenshot_name=f"{element_type}_error")
                self.logger.log.error(f"AssertContainsError > {message}")
                self.task_fail()
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error")
            self.logger.log.error(f"AssertContainsError > Element value of {id} does not contain {expected_value}, but was {actual_value}")
            self.task_fail()

    def assert_success_status(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task:
            self.task
        try:
            smd = self.get_status_msg_dict()
            if smd['messageType'] == "S":
                self.task_pass(msg=f"Status is Success. MsgID & MsgNumber: {smd['messageId']}{smd['messageNumber']}, MsgType: {smd['messageType']}, Msg: {smd['message']} {smd['text']}", ss_name="assert_success_status_pass")
            else:
                self.task_fail(msg=f"Status is not equal 'S'. MsgID & MsgNumber: {smd['messageId']}{smd['messageNumber']}, MsgType: {smd['messageType']}, Msg: {smd['message']} {smd['text']}", ss_name="assert_success_status_fail")
        except Exception as err:
            _msg = f"Unknown error while attempting to check status.|{err}"
            if is_task:
                if fail_on_error:
                    self.task_fail(msg=_msg, ss_name="assert_success_status_fail")
                else:
                    self.task_pass(msg=_msg, ss_name="assert_success_status_pass")
            else:
                self.logger.log.warning(_msg)

    @explicit_wait_before(wait_time=__explicit_wait__)
    def assert_element_present(self, id: str, message: Optional[str] = None, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        _msg = message if message is not None else f"Cannot find element {id}"
        try:
            if not self.is_element(element=id):
                self.on_error(None, _msg, is_task, fail_on_error, ss="assert_element_present_error")
            self.task_pass()
        except Exception as err:
            self.on_error(err, _msg, is_task, fail_on_error, ss="assert_element_present_error")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_string_has_numeric(self, text: str, len_value: Optional[int] = None, message: Optional[str] = None, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> bool:
        if is_task: self.task
        _msg = message if message is not None else ""
        result = False
        try:
            matched_value = re.search("\d+", text).group(0)
        except AttributeError:
            matched_value = None
        except Exception as err:
            self.on_error(err, f"Unknown error while checking if string {text} has numeric value.{_msg}", is_task, fail_on_error, ss="assert_string_has_numeric_error")
        if len_value is not None:
            if matched_value is not None:
                if len(matched_value) == len_value:
                    result = True
        else:
            if matched_value is not None:
                result = True
        if result:
            _msg = message if message is not None else f"String {text} has numeric value: {matched_value}"
            self.task_pass(msg=_msg, ss_name="assert_string_has_numeric_pass")
        else:
            _msg = message if message is not None else f"String {text} has no numeric value"
            self.task_fail(msg=_msg, ss_name="assert_string_has_numeric_fail")
        return result

    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_element_value(self, id: str, expected_value: str, message: Optional[str] = None, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task
        try:
            if self.is_element(element=id):
                actual_value = self.get_value(id=id)
                self.session.findById(id).setfocus()
                self.task_pass()
            if (element_type := self.get_element_type(id)) in text_elements:
                if expected_value != actual_value:
                    _msg = message if message is not None else f"Element value of {id} should be {expected_value}, but was {actual_value}"
                    self.on_error(None, _msg, is_task, fail_on_error, ss=f"{element_type}_error")
            elif element_type in ("GuiCheckBox", "GuiRadioButton"):
                if expected_value := bool(expected_value):
                    if not actual_value:
                        _msg = message if message is not None else f"Element value of {id} should be {expected_value}, but was {actual_value}"
                        self.on_error(None, _msg, is_task, fail_on_error, ss=f"{element_type}_error")
                elif not expected_value:
                    if actual_value:
                        _msg = message if message is not None else f"Element value of {id} should be {expected_value}, but was {actual_value}"
                        self.on_error(None, _msg, is_task, fail_on_error, ss=f"{element_type}_error")
            else:
                _msg = message if message is not None else f"Element value of {id} should be {expected_value}, but was {actual_value}"
                self.on_error(None, _msg, is_task, fail_on_error, ss=f"{element_type}_error")
        except Exception as err:
            _msg = message if message is not None else ""
            self.on_error(err, _msg, is_task, fail_on_error, ss="assert_element_value_error")
    
    assert_element_value_equal = assert_element_value

    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_element_value_not_equal(self, id: str, expected_value: str, message: Optional[str] = None) -> None:
        if self.is_element(element=id):
            actual_value = self.get_value(id=id)
            self.session.findById(id).setfocus()
            self.task_pass()
        if (element_type := self.get_element_type(id)) in text_elements:
            if expected_value == actual_value:
                message = message if message is not None else f"Element value of {id} should not be equal to {expected_value}"
                self.take_screenshot(screenshot_name=f"{element_type}_error")
                self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                self.task_fail()
        elif element_type in ("GuiCheckBox", "GuiRadioButton"):
            if expected_value := bool(expected_value):
                if not actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error")
                    self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                    self.task_fail()
            elif not expected_value:
                if actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error")
                    self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                    self.task_fail()
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error")
            self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
            self.task_fail()


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

    def va01_header(self, sold_to: str, ship_to: str, customer_reference: Optional[str] = None, customer_reference_date: Optional[str] = None, press_enter: Optional[bool] = True) -> None:
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR", text=sold_to)
        self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR", text=ship_to)
        if customer_reference:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=customer_reference)
        else:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=datetime.datetime.now().strftime("%Y%m%d%H%M%S%f"))
        if customer_reference_date:
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK", text=customer_reference_date)
        if press_enter:
            self.sap.send_vkey(vkey="Enter")
        # Handle status msg about duplicate PO values
        result = self.sap.get_status_msg_dict()
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
    
    def va01_handle_apt_item(self, option: Optional[str] = "PROPOSAL") -> None:
        pass
    
    def va01_line_items(self, line_items: list[dict], press_enter: Optional[bool] = True) -> None:
        self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01")
        for item in line_items:
            self.sap.click_element(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POAN")
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,1]", text=item["material"])
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,1]", text=item["target_quantity"])
            self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-VRKME[3,1]", text=item["uom"])
            if "customer_material" in item:
                self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-KDMAT[6,1]", text=item["customer_material"])
            if "item_category" in item:
                self.sap.input_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PSTYV[7,1]", text=item["item_category"])
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



