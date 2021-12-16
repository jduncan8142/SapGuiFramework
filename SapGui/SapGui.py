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


class Gui:
    """
     Python Framework library for controlling the SAP GUI Desktop and focused 
     on testing business processes. The library uses the native SAP GUI scripting engine 
     for interaction with the desktop client application.
    """

    __version__ = '0.0.9'

    def __init__(
        self, 
        test_case: Optional[str] = "My Test Case",
        exit_on_error: Optional[bool] = True, 
        screenshot_on_fail: Optional[bool] = True, 
        screenshot_on_pass: Optional[bool] = True, 
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
        close_sap_on_cleanup: Optional[bool] = True) -> None:
        atexit.register(self.cleanup)
        self.test_case_name: str = test_case
        self.exit_on_error: bool = exit_on_error
        self.screenshot_on_fail: bool = screenshot_on_fail
        self.screenshot_on_pass: bool = screenshot_on_pass
        self.close_sap_on_cleanup: bool = close_sap_on_cleanup
        self.logger: Logger = Logger(log_name=self.test_case_name, log_path=log_path, log_file=log_file, verbosity=verbosity, format=log_format, file_mode=log_file_mode)
        self.__connection_number: int = connection_number
        self.__session_number: int = session_number
        self.explicit_wait: float = explicit_wait
        self.connection_name: str = connection_name if connection_name is not None else ""
        self.sap_gui: win32com.client.CDispatch = None
        self.sap_app: win32com.client.CDispatch = None
        self.connection: win32com.client.CDispatch = None
        self.session: win32com.client.CDispatch = None
        self.screenshot_dir: str = screenshot_dir
        self.monitor: int = int(monitor)
        self.screenshot: Screenshot = Screenshot()
        self.date_format: str = str(date_format)

        if not os.path.exists(self.screenshot_dir):
            self.logger.log.debug(f"Screenshot directory {self.screenshot_dir} does not exist, creating it.")
            try:
                os.makedirs(self.screenshot_dir)
            except Exception as err:
                self.logger.log.error(f"Unable to create screenshot directory {self.screenshot_dir} > {err}")
        self.screenshot.screenshot_directory = self.screenshot_dir
        self.screenshot.monitor = self.monitor

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
        self.passed_tasks: list = []
        self.__task: str = ""
        self.element_id: str = None
        self.element: SapObject = None
    
    def transaction_does_not_exist_strings(self) -> tuple:
        return (
            f"Transactie {self.transaction} bestaat niet", 
            f"Transaction {self.transaction} does not exist", 
            f"Transaktion {self.transaction} existiert nicht")
    
    def cleanup(self):
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
            self.element = self.session.findById(id)
            return True
        except:
            return False

    def take_screenshot(self, screenshot_name: Optional[str] = None, msg: Optional[str] = None) -> None:
        _msg = msg if msg is not None else ""
        _file_names = []
        if not screenshot_name:
            _file_names.append(self.screenshot.shot())
        else:
            _file_names = self.screenshot.shot(name=screenshot_name)
        if _file_names:
            for f in _file_names:
                encoded_img = None
                with open(f, "rb") as f_img:
                    encoded_img = base64.b64encode(f_img.read())
                self.logger.log.shot(f"{_msg}|{f}|{encoded_img}")

    def wait(self, value: Optional[float] = None) -> None:
        """
        Waits for the number of seconds given by value parameter or if value is None the explicit_wait value is used.

        Keyword Arguments:
            value {Optional[float]} -- Number of seconds to wait. (default: {None})
        """
        if value:
            if type(value) is float | int:
                time.sleep(float(value))
            elif type(value) is str:
                _value = None
                try:
                    if "," in value:
                        value = value.replace(",", "").strip()
                    if "." in value:
                        _value = float(value)
                    else:
                        _value = int(value)
                    time.sleep(float(_value))
                except ValueError:
                    self.logger.log.error(f"Unable to convert to int or float for {value} during call to wait, skipping wait timer")
        else:
            time.sleep(self.explicit_wait)

    def fail(self, msg: Optional[str] = None, ss_name: Optional[str] = None) -> None:
        """
        Called from other function responsible for executing tasks. Implements wait if explicit_wait is set and marks the task as PASS.

        Keyword Arguments:
            msg {Optional[str]} -- Message to be logged as error (default: {None})
        """
        self.wait()
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

    def task_passed(self, msg: Optional[str] = None, ss_name: Optional[str] = None) -> None:
        """
        Called from other function responsible for executing test task. Implements wait if explicit_wait is set and marks the task as PASS.

        Keyword Arguments:
            msg {Optional[str]} -- Message to be logged as info (default: {None})
        """
        self.wait()
        if msg:
            self.logger.log.info(msg)
        if self.screenshot_on_pass:
            self.take_screenshot(screenshot_name=ss_name, msg=msg)
        self.task_status = PASS
        self.passed_tasks.append(self.task)

    def get_element_type(self, id: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> str | None:
        """
        Get type information for SAP GUI element.

        Returns:
            SAP GUI Element type -- Type of SAP GUI element.
        """
        _tmp = None
        if is_task:
            self.task()
        try:
            _tmp = self.session.findById(id).type
            if is_task:
                self.task_passed()
        except Exception as err:
            if is_task:
                if fail_on_error: 
                    self.fail(msg=f"Unknown element id: {id} -> {err}")
                else:
                    self.task_passed(msg=f"Unknown element id: {id}|{err}")
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
            ConnectionError: Error while getting SAP session to Windo
            ConnectionError: Unable to get status bar during session connection
            ConnectionError: Unable to get session information
        """
        if is_task:
            self.task()
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
            self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.window}]/sbar")
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
                self.task_passed()
        except Exception as err:
            _msg = f"Unknown error while establishing connection with SAP GUI|{sys.exc_info()[0]}|{err}"
            if is_task:
                if fail_on_error:
                    self.fail(_msg)
                else:
                    self.task_passed(_msg)
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
            self.task()
        if connection_name:
            self.connection_name = connection_name
        try:
            self.connection = self.sap_gui.Children(self.__connection_number)
            if self.connection.Description == self.connection_name:
                self.session = self.connection.children(self.session_number)
                self.wait(2.0)
                self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.session_number}]/wnd[{self.window}]/sbar")
                self.session_info = self.session.info
                self.task_passed()
            else:
                self.fail(msg=f"No existing connection for {self.connection_name} found.", ss_name="connect_to_existing_connection_error")
        except Exception as err:
            _msg = f"Unknown error while trying to establish existing connection for {self.connection_name}|{err}."
            if is_task:
                if fail_on_error:
                    self.fail(msg=_msg, ss_name="connect_to_existing_connection_error")
                else:
                    self.task_passed(msg=_msg, ss_name="connect_to_existing_connection_error")
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
        if is_task:
            self.task()
        if not hasattr(self.sap_app, "OpenConnection"):
            try:
                self.sap_gui = win32com.client.GetObject("SAPGUI")
                if not type(self.sap_gui) == win32com.client.CDispatch:
                    self.fail("Error while getting SAP GUI object using win32com.client")
                self.sap_app = self.sap_gui.GetScriptingEngine
                if not type(self.sap_app) == win32com.client.CDispatch:
                    self.sap_gui = None
                    self.fail("Error while getting SAP scripting engine")
                if connection_name:
                    self.connection_name = connection_name
                self.connection = self.sap_app.OpenConnection(self.connection_name, True)
                self.session = self.connection.children(self.__session_number)
                self.wait(1.0)
                self.sbar = self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.window}]/sbar")
                self.session_info = self.session.info
                self.task_passed(ss_name="open_connection")
            except Exception as err:
                _msg = f"Cannot open connection {self.connection_name}, please check connection name|{err}"
                if is_task:
                    if fail_on_error:
                        self.fail(msg=_msg, ss_name="open_connection")
                    else:
                        self.task_passed(msg=_msg, ss_name="open_connection")
                else:
                    self.logger.log.warning(_msg)

    def get_status_msg_dict(self) -> dict:
        """
        Gets the SAP status message text as a dictonary.

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
    
    def assert_success_status(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task:
            self.task()
        try:
            smd = self.get_status_msg_dict()
            if smd['messageType'] == "S":
                self.task_passed(msg=f"Status is Success. MsgID & MsgNumber: {smd['messageId']}{smd['messageNumber']}, MsgType: {smd['messageType']}, Msg: {smd['message']} {smd['text']}", ss_name="assert_success_status_pass")
            else:
                self.fail(msg=f"Status is not equal 'S'. MsgID & MsgNumber: {smd['messageId']}{smd['messageNumber']}, MsgType: {smd['messageType']}, Msg: {smd['message']} {smd['text']}", ss_name="assert_success_status_fail")
        except Exception as err:
            _msg = f"Unknown error while attempting to check status.|{err}"
            if is_task:
                if fail_on_error:
                    self.fail(msg=_msg, ss_name="assert_success_status_fail")
                else:
                    self.task_passed(msg=_msg, ss_name="assert_success_status_pass")
            else:
                self.logger.log.warning(_msg)

    def exit(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        """
        Exit the current SAP session.

        Keyword Arguments:
            fail_on_error {Optional[bool]} -- If case should fail if there is an error during the execution (default: {True})
            is_task {Optional[bool]} -- If the current function call is a task called by the user (default: {True})
        """
        if is_task:
            self.task()
        try:
            self.connection.closeSession(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]")
            self.connection.closeConnection()
            self.task_passed(msg="Exit successfully.", ss_name="exit")
        except Exception as err:
            _msg = f"Unknown error while attempting to exit SAP session.|{err}"
            if is_task:
                if fail_on_error:
                    self.fail(msg=_msg, ss_name="exit")
                else:
                    self.task_passed(msg=_msg, ss_name="exit")
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
            self.task()
        try:
            self.session.findById(f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.window}]").maximize()
            self.task_passed(msg="Maximize of SAP window successful", ss_name="maximize_window")
        except Exception as err:
            _msg = f"Unknown error while attempting to maximize SAP window.|{err}"
            if is_task:
                if fail_on_error:
                    self.fail(msg=_msg, ss_name="maximize_window")
                else:
                    self.task_passed(msg=_msg, ss_name="maximize_window")
            else:
                self.logger.log.warning(_msg)

    def restart_session(self, connection_name: str, delay: Optional[float] = 0.0, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
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
            self.task()
        try:
            self.connection_name = connection_name if connection_name is not None else self.connection_name
            self.exit()
            self.open_connection(connection_name=self.connection_name)
            self.maximize_window()
            self.wait(value=delay)
            self.task_passed(msg="Successfully restart SAP session.", ss_name="restart_session")
        except Exception as err:
            _msg = f"Unknown error while attempting to restart SAP session.|{err}"
            if is_task:
                if fail_on_error:
                    self.fail(msg=_msg, ss_name="restart_session")
                else:
                    self.task_passed(msg=_msg, ss_name="restart_session")
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
        if is_task:
            self.task()
        t = Timer()
        while not self.is_element(element=id) and t.elapsed() <= timeout:
            self.wait(value=0.5)
        if not self.is_element(element=id):
            if is_task:
                if fail_on_error:
                    self.fail(msg=f"Wait For Element could not find element with id {id}", ss_name="wait_for_element_error")
                else:
                    self.task_passed(msg=f"Wait For Element could not find element with id {id}", ss_name="wait_for_element")
            else:
                self.logger.log.warning(f"Wait For Element could not find element with id {id}")
        else:
            self.task_passed(msg=f"Wait For Element with id {id} successful", ss_name="wait_for_element")

    def get_statusbar_if_error(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> str | None:
        """
        Get SAP statusbar message if statusbar in error state.

        Returns:
            str | None -- Statusbar error text or None
        """
        _tmp = None
        if is_task:
            self.task()
        try:
            if self.sbar.messageType == "E":
                _tmp = f"{self.sbar.findById('pane[0]').text} -> Message no. {self.sbar.messageId.strip('')}:{self.sbar.messageNumber}"
            self.task_passed(msg=f"get_statusbar_if_error was successful", ss_name="get_statusbar_if_error")
        except Exception as err:
            _msg = f"Unhandled error while checking if statusbar had error msg.|{err}"
            if is_task:
                if fail_on_error:
                    self.fail(msg=_msg, ss_name="error_for_get_statusbar_if_error")
                else:
                    self.task_passed(msg=_msg, ss_name="error_for_get_statusbar_if_error")
            else:
                self.log.warning(_msg)
        return _tmp

    def start_transaction(self, transaction: str, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task()
        try:
            self.transaction = transaction.upper()
            self.session.startTransaction(self.transaction)
            self.wait(1.0)
            if (s_msg := str(self.sbar.findById('pane[0]').text).strip(" \n\r\t")) in self.transaction_does_not_exist_strings():
                if fail_on_error:
                    self.fail(msg=f"ValueError|{s_msg}", ss_name="start_transaction_error")
                else:
                    self.task_passed(msg=f"ValueError|{s_msg}", ss_name="start_transaction")
            else:
                self.task_passed(msg=f"Started transaction {self.transaction} successfully|{s_msg}", ss_name="start_transaction")
        except Exception as err:
            _msg = f"Unhandled error during start_transaction|{err}"
            if is_task:
                if fail_on_error:
                    self.fail(msg=_msg, ss_name="start_transaction_error")
                else:
                    self.task_passed(msg=_msg, ss_name="start_transaction_error")
            else:
                self.logger.log.warning(_msg)
    
    start: FunctionType = start_transaction
    Start: FunctionType = start_transaction
    START: FunctionType = start_transaction
    
    def end_transaction(self, fail_on_error: Optional[bool] = True, is_task: Optional[bool] = True) -> None:
        if is_task: self.task()
        try:
            self.session.endTransaction()
            self.task_passed()
        except Exception as err:
            _msg = f"Error ending transaction|{err}"
            if is_task:
                if fail_on_error:
                    self.fail(msg=_msg, ss_name="end_transaction_error")
                else:
                    self.task_passed(msg=_msg, ss_name="end_transaction_error")
            else:
                self.logger.log.warning(_msg)
    
    end: FunctionType = end_transaction
    End: FunctionType = end_transaction
    END: FunctionType = end_transaction
    
    def send_command(self, command: str, fail_on_error: Optional[bool] = True) -> None:
        try:
            self.session.sendCommand(command)
        except Exception as err:
            if fail_on_error:
                self.take_screenshot(screenshot_name="send_command_error")
                self.logger.log.error(f"Error sending command {command}|{err}")
                self.fail()
        self.wait()
        self.task_passed()

    def click_element(self, id: str = None, fail_on_error: Optional[bool] = True) -> None:
        try:
            if (element_type := self.get_element_type(id)) in ("GuiTab", "GuiMenu"):
                self.session.findById(id).select()
            elif element_type == "GuiButton":
                self.session.findById(id).press()
        except Exception as err:
            if fail_on_error:
                self.take_screenshot(screenshot_name="click_element_error")
                self.logger.log.error(f"You cannot use 'Click Element' on element id type {id} > {err}")
                self.fail()
        self.wait()
        self.task_passed()
    
    click = click_element

    def click_toolbar_button(self, table_id: str, button_id: str, fail_on_error: Optional[bool] = True) -> None:
        self.element_should_be_present(table_id)
        try:
            self.session.findById(table_id).pressToolbarButton(button_id)
        except AttributeError:
            self.session.findById(table_id).pressButton(button_id)
        except Exception as err:
            if fail_on_error:
                self.take_screenshot(screenshot_name="click_toolbar_button_error")
                self.logger.log.error(f"Cannot find Table ID/Button ID: {' / '.join([table_id, button_id])}  <-->  {err}")
                self.fail()
        self.wait()
        self.task_passed()

    def doubleclick(self, id: str, item_id: str, column_id: str) -> None:
        if (element_type := self.get_element_type(id)) == "GuiShell":
            self.session.findById(id).doubleClickItem(item_id, column_id)
        else:
            self.take_screenshot(screenshot_name="doubleclick_element_error")
            self.logger.log.error(f"You cannot use 'doubleclick element' on element type {element_type}")
            self.fail()
        self.wait()
        self.task_passed()

    def assert_element_present(self, id: str, message: Optional[str] = None) -> None:
        if not self.is_element(element=id):
            self.take_screenshot(screenshot_name="assert_element_present_error")
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
                    self.take_screenshot(screenshot_name=f"{element_type}_error")
                    self.logger.log.error(f"AssertEqualError > Element value of {id} should be {expected_value}, but was {actual_value}")
                    self.fail()
            elif not expected_value:
                if actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error")
                    self.logger.log.error(f"AssertEqualError > Element value of {id} should be {expected_value}, but was {actual_value}")
                    self.fail()
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error")
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
                self.take_screenshot(screenshot_name=f"{element_type}_error")
                self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                self.fail()
        elif element_type in ("GuiCheckBox", "GuiRadioButton"):
            if expected_value := bool(expected_value):
                if not actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error")
                    self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                    self.fail()
            elif not expected_value:
                if actual_value:
                    self.take_screenshot(screenshot_name=f"{element_type}_error")
                    self.logger.log.error(f"AssertNotEqualError > Element value of {id} should not be equal to {expected_value}")
                    self.fail()
        else:
            self.take_screenshot(screenshot_name=f"{element_type}_error")
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
                self.take_screenshot(screenshot_name=f"{element_type}_error")
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

    def get_value(self, id: str, exit_on_error: Optional[bool] = True) -> Any:
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
                return _value
            else:
                return None
        except Exception as err:
            self.take_screenshot(screenshot_name="get_value_error.jpg")
            self.logger.log.error(f"Cannot get value for element type {element_type} for id {id} -> {err}")
            self.fail(exit_on_error=exit_on_error)

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

    def string_generator(self, size: Optional[int]=6, chars: Optional[str]=string.ascii_uppercase + string.digits) -> str:
        selected_chars = []
        for i in range(size):
            selected_chars.append(random.choice(chars))
        return ''.join(selected_chars)

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
    
    def back(self) -> None:
        self.send_vkey(vkey="F3")


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



