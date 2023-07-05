from dotenv import load_dotenv
from typing import Any, Optional
import win32com.client
from Flow.Data import Case, TextElements, VKEYS, Table, BrowserType
from Flow.Results import Result
from Flow.Actions import Step
from Logging.Logging import Logger, LoggingConfig
from Core.Utilities import *
from Core.SAP import *
from time import sleep
import atexit
import base64
import datetime
import re
import json
import os
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import chromedriver_binary


class Session:
    __version__: str = "0.1.4"
    __explicit_wait__: float = 0.0
    
    def __init__(self) -> None:
        """
        1. Load any available .env file
        2. Create Case instance 
        3. Create Logger instance
        4. Set global __explicit_wait__ variable 
        5. Create Session variables and set defaults
        6. Register Session.cleanup function so it is called on exit
        """
        load_dotenv()
        self.case: Case = Case()
        self.logger: Logger|None = None
        self.logger = Logger(config=self.case.LogConfig)
        Session.__explicit_wait__ = self.case.ExplicitWait
        self.web_driver: webdriver|None = None
        self.web_element: WebElement|None = None
        self.web_iframe: WebElement|None = None
        self.web_wait: float = float(os.getenv("HTML_WAIT", "3.0"))
        self.__connection_number: int = 0
        self.__session_number: int = 0
        self.__window_number: int = 0
        self.connection_name: str|None = None
        self.sap_gui: win32com.client.CDispatch|None = None
        self.sap_app: win32com.client.CDispatch|None = None
        self.connection: win32com.client.CDispatch|None = None
        self.session: win32com.client.CDispatch|None = None
        self.session_info: win32com.client.CDispatch|None = None
        self.main_window: win32com.client.CDispatch|None = None
        self.mbar: win32com.client.CDispatch|None = None
        self.tbar0: win32com.client.CDispatch|None = None
        self.titl: win32com.client.CDispatch|None = None
        self.tbar1: win32com.client.CDispatch|None = None
        self.usr: win32com.client.CDispatch|None = None
        self.sbar: win32com.client.CDispatch|None = None
        self.current_element: win32com.client.CDispatch|None = None
        self.current_transaction: str|None = None
        self.current_step: Step|None = None
        atexit.register(self.cleanup)
    
    def __post_init__(self) -> None:
        """
        Post init function of Session class. 
        If user did not pass in a class<Step> object for Session.current_step, create a new initial step object with default values.
        """
        if self.current_step is None:
            self.current_step = Step(
                Action="Create Session", 
                ElementId="", 
                Args=[],
                Name="Create New Session", 
                Description="Creates and return a new SAP session object.")
    
    # Screenshot Actions
    def hard_copy(self, filename: str, image_type: Optional[str] = "PNG", pos: Optional[tuple[int, int, int, int]] = None) -> bytes|None:
        """
        Capture SAP GUI screenshot

        Arguments:
            filename {str} -- Filename used to save file from main_window.HandCopy function of SAP GUI

        Keyword Arguments:
            image_type {Optional[str]} -- Type of image file (default: {"PNG"})
            pos {Optional[tuple[int, int, int, int]]} -- Optional tuple of (left, top, width, height) (default: {None})

        Returns:
            bytes|None -- Returns the bytes of the screenshot or None
        """
        shot_bytes: bytes|None = None
        try:
            if pos is not None:
                img = self.main_window.HardCopy(
                    filename, 
                    image_type, 
                    pos[0], 
                    pos[1], 
                    pos[2], 
                    pos[3])
            else:
                img = self.main_window.HardCopy(filename, image_type)
                with open(img, "rb") as f_img:
                    shot_bytes = base64.b64encode(f_img.read())
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during hard_copy", 
                ss_name="hard_copy_exception", 
                error=err)
        return shot_bytes

    @explicit_wait_before(wait_time=__explicit_wait__)
    def capture_fullscreen(self, screenshot_name: str) -> bytes|None:
        """
        Captures a screenshot of the SAP GUI main window as bytes.
        Bytes can then be saved to a file or base64 encoded and added to the log.

        Arguments:
            screenshot_name {str} -- Name of screenshot

        Returns:
            bytes|None -- Returns the bytes of the screenshot or None
        """
        shot_bytes: bytes|None = None
        try:
            shot_bytes = self.hard_copy(screenshot_name, "PNG")
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during screen capture", 
                ss_name="capture_fullscreen_exception", 
                error=err)
        return shot_bytes
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def capture_region(self, screenshot_name: str, pos: tuple[int, int, int, int]) -> bytes|None:
        """
        Captures a screenshot of a provide SAP GUI window region as bytes.
        Bytes can then be saved to a file or base64 encoded and added to the log.

        Arguments:
            screenshot_name {str} -- Name of screenshot
            pos {tuple[int, int, int, int]} -- SAP GUI window region to capture

        Returns:
            bytes|None -- Returns the bytes of the screenshot or None
        """
        shot_bytes: bytes|None = None
        try:
            shot_bytes = self.hard_copy(screenshot_name, "PNG", pos)
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during region capture", 
                ss_name="capture_region_exception", 
                error=err)
        return shot_bytes
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def capture_element(self, element_id: str, screenshot_name: Optional[str|None] = None) -> bytes|None:
        """
        Captures a screenshot of a provide SAP GUI element as bytes. 
        Bytes can then be saved to a file or base64 encoded and added to the log.

        Arguments:
            element_id {str} -- SAP GUI element to capture
            screenshot_name {Optional[str]} -- Name of screenshot

        Returns:
            bytes|None -- Returns the bytes of the screenshot or None
        """
        shot_bytes: bytes|None = None
        try:
            __element = self.session.FindById(self.ace_id(element_id))
            __pos = (
                __element.ScreenLeft, 
                __element.ScreenTop, 
                __element.Width, 
                __element.Height
            )
            if screenshot_name is not None:
                shot_bytes = self.hard_copy(screenshot_name, "PNG", __pos)
            else:
                shot_bytes = self.hard_copy(__element.Name, "PNG", __pos)
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during element capture", 
                ss_name="capture_element_exception", 
                error=err)
        return shot_bytes
    
    # Helpers
    def get_env(self, var: str) -> str:
        """
        Gets a provided user/system environment variable. 
        Useful to fetch usernames/password or other sensitive data that you might not want to store in a public data file.

        Arguments:
            var {str} -- Name of environment variable to get

        Returns:
            Any|None -- Returns the value of the environment variable or None
        """
        return str(os.getenv(var), "")
    
    def is_element(self, element: str) -> bool:
        """
        Handler to check is a provided string is a valid SAP GUI element id for the current window and session.

        Arguments:
            element {str} -- SAP GUI element id

        Returns:
            bool -- Returns True if element exist otherwise False
        """
        try:
            __element = self.ace_id(element)
            self.current_element = self.session.findById(__element)
            self.step_pass(
                msg="Element: %s is valid" % __element, 
                ss_name="is_element_pass")
            return True
        except Exception as err:
            self.handle_unknown_exception(
                msg="SAP element %s id not found." % __element, 
                ss_name="is_element_exception",
                error=err)
        return False

    def exit(self) -> None:
        """
        Exits the current SAP GUI session and window.
        """
        try:
            self.connection.closeSession(self.ace_id())
            self.connection.closeConnection()
            self.step_pass(
                msg=f"Successfully exited session.", 
                ss_name="exit_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unknown exception while exiting session.",
                ss_name="exit_exception", 
                error=err)
    
    def cleanup(self) -> None:
        """
        Handler for cleanup at end of session usage. 
        If CloseOnCleanup flag is True then makes sure to exit open sessions even if process is terminated due to error.
        Updates Case.Status.Result to Result.FAIL is any step has fails else to Result.PASS. 
        Logs to case's final result status.
        """
        if self.case.CloseSAPOnCleanup:
            self.exit()
        if self.case.Status.Result is None:
            if len(self.case.Status.FailedSteps) != 0:
                self.case.Status.Result = Result.FAIL
        else:
            self.case.Status.Result = Result.PASS
        self.documentation(
            f"{self.case.Name} completed with status: {self.case.Status.Result.value}")

    def wait(self, seconds: float) -> None:
        """
        Wait/Sleep for a given number fo seconds.
        This function logs the wait in the current case's log. 
        This function is blocking as it uses the standard python time.sleep function.

        Arguments:
            seconds {float} -- Number of seconds to wait
        """
        if seconds == 1.0:
            self.documentation(f"Waiting 1 second...")
        else:
            self.documentation(f"Waiting {seconds} seconds...")
        sleep(seconds)
    
    def wait_for_element(self, id: str, timeout: Optional[float] = 60.0) -> None:
        """
        Waits <timeout> seconds for the provided SAP GUI element to become ready or available.
        Rechecks ever half second.

        Arguments:
            id {str} -- SAP GUI element id

        Keyword Arguments:
            timeout {Optional[float]} -- timeout in seconds to wait for element to become available (default: {60.0})
        """
        try:
            __id = self.ace_id(id)
            t = Timer()
            while True:
                if not self.is_element(element=__id) and t.elapsed() <= timeout:
                    self.wait(seconds=0.5)
                else:
                    break
            if not self.is_element(element=__id):
                self.step_fail(
                    msg=f"No element found with id: {__id}", 
                    ss_name="wait_for_element_fail")
            else:
                self.step_pass(
                    msg=f"Found element with id: {__id}", 
                    ss_name="wait_for_element_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"Unhandled exception waiting for element id: {id}", 
                ss_name="wait_for_element_exception", 
                error=err)

    def try_and_continue(self, func: object, *args, **kwargs) -> Any:
        """
        Wrapper to try running provided function and continue logging a PASS status even if there is an error or failure.

        Arguments:
            func {object} -- Function to execute

        Returns:
            Any -- Returns the result of the called function if function has a return value or None
        """
        __result = None
        try:
            if hasattr(self, func) and callable(func := getattr(self, func)):
                __result = func(*args, **kwargs)
        except Exception as err:
            self.logger.log.info(f"Unhandled exception during Try and Continue wrapped function: {func}")
            self.current_step.Status.Result = Result.WARN
            self.current_step.Status.Error = err
            self.case.Status.PassedSteps.append(self.current_step)
            if self.case.ScreenShotOnPass:
                self.case.Status.PassedScreenShots.append(self.capture_fullscreen(screenshot_name="try_and_continue_exception"))            
        return __result
    
    def parse_document_number(self) -> str|None:
        """
        Extracts and return a document number for the SAP GUI status bar element. 
        Often used immediately after the save function to get the newly created SAP document, 
        shown in the status bar.

        Returns:
            str -- Parsed document number from status bar
        """
        return re.search("\d+", self.sbar.Text).group(0)

    def ace_id(self, id: Optional[str] = None) -> str:
        """
        Handler to auto complete partial SAP GUI element ids.
        If not id argument is provided the SAP GUI element id of the main window of the current session is returned.

        Keyword Arguments:
            id {Optional[str]} -- Optional (partial) SAP GUI element id complete (default: {None})

        Returns:
            str -- Full SAP GUI element id
        """
        base_id: str = f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/wnd[{self.__window_number}]"
        if id in ("",  " ", None):
            return base_id
        elif id.startswith("usr"):
            return f"{base_id}/{id}"
        elif id.startswith("/usr"):
            return f"{base_id}{id}"
        elif id.startswith("wnd"):
            return f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]/{id}"
        elif id.startswith("/wnd"):
            return f"/app/con[{self.__connection_number}]/ses[{self.__session_number}]{id}"
        elif id.startswith("ses"):
            return f"/app/con[{self.__connection_number}]/{id}"
        elif id.startswith("/ses"):
            return f"/app/con[{self.__connection_number}]{id}"
        elif id.startswith("con"):
            return f"/app/{id}"
        elif id.startswith("/con"):
            return f"/app{id}"
        elif id.startswith("app"):
            return f"/{id}"
        elif id.startswith("/app"):
            return id
        else:
            return id
    
    def documentation(self, msg: Optional[str] = None) -> None:
        """
        Handler for documenting a msg in the log.

        Keyword Arguments:
            msg {Optional[str]} -- Optional msg to be logged (default: {f"{self.current_step.Name} -- {self.current_step.Description}"})
        """
        _msg = msg if msg is not None else f"{self.current_step.Name} -- {self.current_step.Description}"
        if _msg is not None and _msg != "" and _msg != "--":
            self.logger.log.documentation(_msg)
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def step_fail(self, msg: Optional[str] = None, ss_name: Optional[str] = None, error: Optional[str] = None) -> None:
        """
        Handler for test case steps, called when step fails or is unsuccessful.
        Does not ensure step result only that step returned a python error.

        Keyword Arguments:
            msg {Optional[str]} -- Optional msg to include in logging (default: {None})
            ss_name {Optional[str]} -- Name of screenshot if ScreenshotOnFail flag is set for the current case (default: {None})
            error {Optional[str]} -- Optional passing of standard python error text to be logged (default: {None})
        """
        if msg:
            self.logger.log.error(msg)
        self.current_step.Status.Result = Result.FAIL
        self.current_step.Status.Error = error if error is not None else ""
        self.case.Status.FailedSteps.append(self.current_step)
        if self.case.ScreenShotOnFail:
            __ss_name = ss_name if ss_name is not None else f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
            self.case.Status.FailedScreenShots.append(self.capture_fullscreen(screenshot_name=__ss_name))
        if self.case.ExitOnFail:
            sys.exit()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def step_pass(self, msg: Optional[str] = None, ss_name: Optional[str] = None) -> None:
        """
        Handler for test case steps, called when step passes or is successful.
        Does not ensure step result only that step completed without throwing a python error.

        Keyword Arguments:
            msg {Optional[str]} -- Optional msg to include in logging (default: {None})
            ss_name {Optional[str]} -- Name of screenshot if ScreenshotOnPass flag is set for the current case (default: {None})
        """
        if msg:
            self.logger.log.info(msg)
        self.current_step.Status.Result = Result.PASS
        self.case.Status.PassedSteps.append(self.current_step)
        if self.case.ScreenShotOnPass:
            __ss_name = ss_name if ss_name is not None else f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
            self.case.Status.PassedScreenShots.append(self.capture_fullscreen(screenshot_name=__ss_name))
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def handle_unknown_exception(self, msg: Optional[str] = None, ss_name: Optional[str] = None, error: Optional[str] = None) -> None:
        """
        Handler for python errors during execution of test steps. 
        If the current case's FailOnError flag is set to True the case is marked as failed, 
        else if the flag is False the error is logged as a warning msg and test case execution continues.

        Keyword Arguments:
            msg {Optional[str]} -- Optionals msg used when calling step_fail when FailOnError flag is True (default: {None})
            ss_name {Optional[str]} -- Optionals name of screenshot used when calling step_fail when FailOnError flag is True (default: {None})
            error {Optional[str]} -- Optional passing of standard python error text to be logged (default: {None})
        """
        if self.case.FailOnError:
            self.step_fail(msg=msg, ss_name=ss_name, error=error)
        else:
            self.logger.log.warning(msg=msg)
    
    def new_case(
        self, 
        name: Optional[str] = None, 
        desc: Optional[str] = None, 
        bpo: Optional[str] = None, 
        ito: Optional[str] = None, 
        doc_link: Optional[str] = None, 
        case_path: Optional[Path] = None, 
        log_config: Optional[LoggingConfig] = None, 
        date_format: Optional[str] = None, 
        explicit_wait: Optional[float] = None, 
        screenshot_on_pass: Optional[bool] = None, 
        screenshot_on_fail: Optional[bool] = None, 
        fail_on_error: Optional[bool] = None, 
        exit_on_fail: Optional[bool] = None, 
        close_on_cleanup: Optional[bool] = None, 
        system: Optional[str] = None, 
        steps: Optional[list[Step]] = None, 
        data: Optional[dict] = None) -> None:
        """
        Create a new test case instance.

        Keyword Arguments:
            name {Optional[str]} -- Name of test case (default: {f"test_{datetime.now().strftime('%m%d%Y_%H%M%S')}"})
            desc {Optional[str]} -- Description of test case (default: {""})
            bpo {Optional[str]} -- Business Process Owner or responsible business user (default: {"Business Process Owner"})
            ito {Optional[str]} -- IT Owner or responsible IT user (default: {"Technical Owner"})
            doc_link {Optional[str]} -- Link to external documentation (default: {""})
            case_path {Optional[Path]} -- External path for test case data (default: {function<get_main_dir>})
            log_config {Optional[LoggingConfig]} -- Custom log configuration object (default: {class<LoggingConfig>})
            date_format {Optional[str]} -- Custom string date format (default: {"%m/%d/%Y"})
            explicit_wait {Optional[float]} -- Sets the wait time in seconds (default: {0.25})
            screenshot_on_pass {Optional[bool]} -- Sets flag to capture screenshots when steps pass (default: {False})
            screenshot_on_fail {Optional[bool]} -- Sets flag to capture screenshots when steps fail (default: {False})
            fail_on_error {Optional[bool]} -- Sets flag to fail steps if there is a python error (default: {True})
            exit_on_fail {Optional[bool]} -- Sets flag to stop and exit the script if there is a failed step (default: {True})
            close_on_cleanup {Optional[bool]} -- Sets flag to close the open SAP GUI window/session during cleanup (default: {True})
            system {Optional[str]} -- _description_ (default: {""})
            steps {Optional[list[Step]]} -- _description_ (default: {class<list>})
            data {Optional[dict]} -- _description_ (default: {class<object>})
        """
        __name = name if name is not None else Case.default_name()
        __desc = desc if desc is not None else Case.empty_string()
        __bpo = bpo if bpo is not None else Case.default_business_process_owner()
        __ito = ito if ito is not None else Case.default_it_owner()
        __doc_link = doc_link if doc_link is not None else Case.empty_string()
        __case_path = case_path if case_path is not None else Case.default_case_path()
        __log_config = log_config if log_config is not None else Case.default_log_config()
        __date_format = date_format if date_format is not None else Case.default_date_format()
        __explicit_wait = explicit_wait if explicit_wait is not None else Case.default_explicit_wait()
        __screenshot_on_pass = screenshot_on_pass if screenshot_on_pass is not None else Case.ScreenShotOnPass
        __screenshot_on_fail = screenshot_on_fail if screenshot_on_fail is not None else Case.ScreenShotOnFail
        __fail_on_error = fail_on_error if fail_on_error is not None else Case.FailOnError
        __exit_on_fail = exit_on_fail if exit_on_fail is not None else Case.ExitOnFail
        __close_on_cleanup = close_on_cleanup if close_on_cleanup is not None else Case.CloseSAPOnCleanup
        __system = system if system is not None else Case.default_system()
        __steps = steps if steps is not None else Case.empty_list_factory()
        __data = data
        self.case = Case(
            Name = __name, 
            Description = __desc, 
            BusinessProcessOwner = __bpo, 
            ITOwner = __ito, 
            DocumentationLink = __doc_link, 
            CasePath = __case_path, 
            LogConfig = __log_config, 
            DateFormat = __date_format, 
            ExplicitWait = __explicit_wait, 
            ScreenShotOnPass = __screenshot_on_pass, 
            ScreenShotOnFail = __screenshot_on_fail, 
            ExitOnFail = __exit_on_fail, 
            FailOnError = __fail_on_error, 
            CloseSAPOnCleanup = __close_on_cleanup, 
            System = __system, 
            Steps = __steps, 
            Data = __data)
        self.collect_case_meta_data()
    
    def new_step(self, action: str, id: Optional[str] = "", name: Optional[str] = None, desc: Optional[str] = None, *args, **kwargs) -> None:
        """
        Create a new test case step and add it to the current Case's step list.

        Arguments:
            action {str} -- calling functions name

        Keyword Arguments:
            id {Optional[str]} -- SAP GUI element id, if used (default: {""})
            name {Optional[str]} -- Optional step name for reporting (default: {None})
            desc {Optional[str]} -- Optional step description for reporting (default: {None})
        """
        try:
            __action = action
            __name = name if name is not None else __action.replace("_", " ").title()
            __desc = desc if desc is not None else ""
            self.current_step = Step(
                Action = __action, 
                ElementId = id, 
                Args = args, 
                Kwargs = kwargs,
                Name = __name, 
                Description = __desc)
            self.collect_step_meta_data()
            self.case.Steps.append(self.current_step)
        except Exception as err:
            self.logger.log.warning(msg=f"Unhandled exception while creating new step|{err}")

    @explicit_wait_before(wait_time=__explicit_wait__)
    def collect_step_meta_data(self) -> None:
        """
        Collect SAP metadata for current test step.
        Data collected includes:
        - Application Server
        - Language
        - Program
        - Response Time
        - Round Trips
        - Screen Number
        - System Name
        - System Number
        - System Session Id
        - Transaction
        - User Id
        """
        try:
            if self.current_step and self.session:
                self.current_step.ApplicationServer = self.session_info.ApplicationServer
                self.current_step.Language = self.session_info.Language
                self.current_step.Program = self.session_info.Program
                self.current_step.ResponseTime = self.session_info.ResponseTime
                self.current_step.RoundTrips = self.session_info.RoundTrips
                self.current_step.ScreenNumber = self.session_info.ScreenNumber
                self.current_step.SystemName = self.session_info.SystemName
                self.current_step.SystemNumber = self.session_info.SystemNumber
                self.current_step.SystemSessionId = self.session_info.SystemSessionId
                self.current_step.Transaction = self.session_info.Transaction
                self.current_step.User = self.session_info.User
        except Exception as err:
            self.logger.log.warning(msg=f"Unhandled exception while collecting step metadata|{err}")
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def collect_case_meta_data(self) -> None:
        """
        Collects SAP version metadata of current session.
        """
        try:
            if self.case and self.session:
                self.case.SapMajorVersion = self.sap_app.MajorVersion
                self.case.SapMinorVersion = self.sap_app.MinorVersion
                self.case.SapPatchLevel = self.sap_app.PatchLevel
                self.case.SapRevision = self.sap_app.Revision
        except Exception as err:
            self.logger.log.warning(msg=f"Unhandled exception while collecting case metadata|{err}")
    
    def load_case_from_json_file(self, data_file: str) -> dict:
        """
        Load test case data from a json file

        Arguments:
            data_file {str} -- Path the json data file.

        Returns:
            dict -- Returns a python dictionary of the parsed json data.
        """
        __data: dict = json.load(open(data_file, "rb"))
        self.case.Name = __data.get("case_name", f"test_{datetime.datetime.now().strftime('%m%d%Y_%H%M%S')}")
        self.case.Description = __data.get("description", "")
        self.case.BusinessProcessOwner = __data.get("business_owner", "Business Process Owner")
        self.case.ITOwner = __data.get("it_owner", "Technical Owner")
        self.case.DocumentationLink = __data.get("doc_link", "")
        self.case.CasePath = __data.get("case_path", "")
        self.case.DateFormat = __data.get("date_format", "%m/%d/%Y")
        self.case.ExplicitWait = __data.get("explicit_wait", 0.25)
        self.case.ScreenShotOnPass = __data.get("screenshot_on_pass", False)
        self.case.ScreenShotOnFail = __data.get("screenshot_on_fail", False)
        self.case.FailOnError = __data.get('fail_on_error', True)
        self.case.ExitOnFail = __data.get("exit_on_fail", True)
        self.case.CloseSAPOnCleanup = __data.get("close_sap_on_cleanup", True)
        self.case.System = __data.get("system", "")
        self.case.Data = __data
        return self.case.Data
    
    # Connection Actions
    def open_connection(self, connection_name: str) -> None:
        """
        Open win32 api connection to SAP GUI application.

        Arguments:
            connection_name {str} -- SAP environment name to connect with, can be found in the login pad.
        """
        self.new_step(action="open_connection", connection_name=connection_name)
        self.connection_name = connection_name if connection_name else self.connection_name
        self.documentation(msg=f"Opening connection for {self.connection_name}")
        if not hasattr(self.sap_app, "OpenConnection"):
            try:
                self.sap_gui = win32com.client.GetObject("SAPGUI")
                if not type(self.sap_gui) == win32com.client.CDispatch:
                    self.step_fail("Error while getting SAP GUI object using win32com.client")
                self.sap_app = self.sap_gui.GetScriptingEngine
                if not type(self.sap_app) == win32com.client.CDispatch:
                    self.sap_gui = None
                    self.step_fail("Error while getting SAP scripting engine")
                __conns = self.sap_app.connections
                if len(__conns) == 0:
                    self.connection = self.sap_app.OpenConnection(self.connection_name, True)
                    self.__connection_number = self.connection.Id[-2]
                else:
                    for conn in __conns:
                        if conn.description == connection_name:
                            self.connection = conn
                            self.__connection_number = self.connection.Id[-2]
                    if self.connection is None:
                        self.connection = self.sap_app.OpenConnection(self.connection_name, True)
                        self.__connection_number = self.connection.Id[-2]
                __sessions = self.connection.sessions
                if len(__sessions) == 0:
                    self.session = self.connection.children(self.__session_number)
                    self.__session_number = self.session.Id[-2]
                else:
                    self.session = __sessions[0]
                    self.__session_number = self.session.Id[-2]
                self.collect_session_info()
                self.step_pass(
                    msg=f"Connection open for {self.connection_name}", 
                    ss_name="open_connection_success")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while open connection {self.connection_name}, \
                        check connection name", 
                    ss_name="open_connection_exception", 
                    error=err) 

    # Session Actions
    def restart_session(self, delay: Optional[float] = 0.0) -> None:
        """
        Restart the SAP GUI session

        Keyword Arguments:
            delay {Optional[float]} -- Time in seconds to wait after relaunching the SAP GUI session. (default: {0.0})
        """
        self.new_step(action="restart_session")
        try:
            self.exit()
            self.open_connection(self.connection_name)
            self.wait(seconds=delay)
            self.maximize_window()
            self.step_pass(
                msg="Restarted session successfully.", 
                ss_name="restart_session_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception while restarting session.", 
                ss_name="restart_session_exception", 
                error=err)

    @explicit_wait_before(wait_time=__explicit_wait__)
    def collect_session_info(self) -> None:
        """
        Collects various data point for the current SAP GUI session and SAP window elements.
        """
        try:
            if self.session:
                self.wait_for_element(self.ace_id())
                self.main_window = self.session.findById(self.ace_id())
                self.__window_number = self.main_window.Id[-2]
                self.mbar = self.session.findById(f"{self.ace_id()}/mbar")
                self.tbar0 = self.session.findById(f"{self.ace_id()}/tbar[0]")
                self.titl = self.session.findById(f"{self.ace_id()}/titl")
                self.tbar1 = self.session.findById(f"{self.ace_id()}/tbar[1]")
                self.usr = self.session.findById(f"{self.ace_id()}/usr")
                self.sbar = self.session.findById(f"{self.ace_id()}/sbar")
                self.session_info = self.session.info
        except Exception as err:
            self.logger.log.warning(msg=f"Unhandled exception while collecting session info|{err}")
    
    # Window Actions
    @explicit_wait_before(wait_time=__explicit_wait__)
    def check_for_modal(self, match_text: str, match_id: Optional[str] = None, is_match: Optional[bool]=True, action: Optional[object]=None, **kwargs) -> bool:
        """
        Check for open SAP modal windows (popup windows or msg box windows).

        Arguments:
            match_text {str} -- Text to match against the modal window's title text

        Keyword Arguments:
            match_id {Optional[str]} -- Optional SAP GUI element id to match against rather than the modal window title element (default: {None})
            is_match {Optional[bool]} -- Optional boolean flag to invert the result of the element text match (default: {True})
            action {Optional[object]} -- Optional function to execute on successful match (default: {None})

        Returns:
            bool -- Returns True if there is a modal type window open otherwise False
        """
        self.new_step(action="check_for_modal", match_text=match_text, match_id=match_id, is_match=is_match, my_action=action, **kwargs)
        modal_window = None
        try:
            modal_window = self.session.ActiveWindow
        except Exception as err:
            self.handle_unknown_exception(
                f"Unable to check for popup.", 
                ss_name="check_for_modal_exception", 
                error=err)
        if modal_window is not None:
            if modal_window.Type == "GuiModalWindow":
                if match_id is None:
                    if match_text in modal_window.Text:
                        if is_match:
                            if action is not None:
                                action(**kwargs)
                                return True
                            else:
                                return True
                    else:
                        if not is_match:
                            if action is not None:
                                action(**kwargs)
                                return True
                            else:
                                return True
                else:
                    if match_id is not None:
                        __text = None
                        try:
                            __text = self.session.FindById(match_id).Text
                        except Exception as err:
                            self.handle_unknown_exception(
                                f"Unable to locate match_id: {match_id}.", 
                                ss_name="check_for_modal_match_id_exception", 
                                error=err)
                        if match_text in __text:
                            if is_match:
                                if action is not None:
                                    action(**kwargs)
                                    return True
                                else:
                                    return True
                        else:
                            if not is_match:
                                if action is not None:
                                    action(**kwargs)
                                    return True
                                else:
                                    return True
        return False
    
    def start_transaction(self, transaction: str) -> None:
        """
        Starts a provided transaction in the current SAP GUI session.

        Arguments:
            transaction {str} -- A valid SAP transaction
        """
        self.new_step(action="start_transaction", transaction=transaction)
        self.current_transaction = transaction.upper()
        try:
            self.session.startTransaction(self.current_transaction)
            self.step_pass(
                msg=f"Successfully started transaction: {self.current_transaction}", 
                ss_name="start_transaction_pass")
        except Exception as err:
            self.handle_unknown_exception(
                f"Unable to start transaction: {self.current_transaction}", 
                ss_name="start_transaction_exception", 
                error=err)
    
    def end_transaction(self) -> None:
        """
        Ends the currently active SAP transaction and returns the current session to the Easy Access Menu.
        """
        self.new_step(action="end_transaction", transaction=self.current_transaction)
        try:
            self.session.endTransaction()
            self.step_pass(
                msg=f"Successfully ended transaction: {self.current_transaction}", 
                ss_name="end_transaction_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during end_transaction.", 
                ss_name="end_transaction_exception", 
                error=err)
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def set_v_scrollbar(self, id: str, pos: int) -> None:
        """
        Set the vertical scroll position of the provided SAP GUI scrollbar.

        Arguments:
            id {str} -- SAP GUI element id of the vertical scrollbar
            pos {int} -- Integer value to set the scroolbar position
        """
        self.new_step(action="set_v_scrollbar", id=id, pos=pos)
        if self.is_element(id):
            try:
                self.current_element.verticalScrollbar.position = pos
                self.step_pass(
                    msg=f"Successfully set scrollbar: {self.current_element.Id} \
                        to position: {pos}.", 
                    ss_name="set_v_scrollbar_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    f"Unable to set vertical scrollbar: {self.current_element.Id} \
                        to position: {pos}", 
                    ss_name="set_v_scrollbar_exception",
                    error=err)

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_v_scrollbar(self, id: str) -> int|None:
        """
        Get the current vertical position of the provided SAP GUI scrollbar 

        Arguments:
            id {str} -- SAP GUI element id of the vertical scrollbar

        Returns:
            int|None -- Returns an integer value or None
        """
        self.new_step(action="get_v_scrollbar", id=id)
        __position: int = None
        if self.is_element(id):
            try:
                __position = self.current_element.verticalScrollbar.position
                self.step_pass(
                    msg=f"Successfully got position of vertical scrollbar: {self.current_element.Id}.", 
                    ss_name="get_v_scrollbar_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    f"Unable to get position of vertical scrollbar: {self.current_element.Id}", 
                    ss_name="get_v_scrollbar_exception",
                    error=err)
        return __position

    @explicit_wait_before(wait_time=__explicit_wait__)
    def set_h_scrollbar(self, id: str, pos: int) -> None:
        """
        Set the horizontal scroll position of the provided SAP GUI scrollbar. 

        Arguments:
            id {str} -- SAP GUI element id of the horizontal scrollbar
            pos {int} -- Integer value to set the scroolbar position
        """
        self.new_step(action="set_h_scrollbar", id=id, pos=pos)
        if self.is_element(id):
            try:
                self.current_element.horizontalScrollbar.position = pos
                self.step_pass(
                    msg=f"Successfully set horizontal scrollbar: {self.current_element.Id} \
                        to position: {pos}.", 
                    ss_name="set_h_scrollbar_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    f"Unable to set horizontal scrollbar: {self.current_element.Id} \
                        to position: {pos}", 
                    ss_name="set_h_scrollbar_exception",
                    error=err)

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_h_scrollbar(self, id: str) -> int|None:
        """
        Get the current horizontal position of the provided SAP GUI scrollbar

        Arguments:
            id {str} -- SAP GUI element id of the horizontal scrollbar

        Returns:
            int|None -- Returns an integer value or None
        """
        self.new_step(action="get_h_scrollbar", id=id)
        __position: int = None
        if self.is_element(id):
            try:
                __position = self.current_element.horizontalScrollbar.position
                self.step_pass(
                    msg=f"Successfully got position of horizontal scrollbar: {self.current_element.Id}.", 
                    ss_name="get_h_scrollbar_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    f"Unable to get position of horizontal scrollbar: {self.current_element.Id}", 
                    ss_name="get_h_scrollbar_exception",
                    error=err)
        return __position

    @explicit_wait_before(wait_time=__explicit_wait__)
    def maximize_window(self) -> None:
        """
        Maximizes the current SAP GUI window.
        """
        self.new_step(action="maximize_window")
        try:
            self.main_window.maximize()
            self.step_pass(msg="Window maximized.", ss_name="maximize_window_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception maximizing window.", 
                ss_name="maximize_window_exception", 
                error=err)

    # Keyboard & Mouse Actions
    @explicit_wait_after(wait_time=__explicit_wait__)
    def click_element(self, id: str) -> None:
        """
        Clicks the SAP GUI element identified by the id provided. 

        Arguments:
            id {str} -- Id of the SAP GUI element to be clicked
        """
        self.new_step(action="click_element", id=id)
        if self.is_element(id):
            try:
                if self.current_element.Type in ("GuiTab", "GuiMenu", "GuiRadioButton"):
                    self.current_element.Select()
                    self.step_pass(msg=f"Successfully clicked element: {id}", ss_name="click_element_success")
                elif self.current_element.Type == "GuiButton":
                    self.current_element.Press()
                    self.step_pass(msg=f"Successfully clicking GuiButton: {id}", ss_name="click_gui_button_success")
                else:
                    self.step_fail(msg=f"Unable to click element: {id}", ss_name="click_element_failed")
            except Exception as err:
                self.handle_unknown_exception(msg=f"Unhandled exception while clicking element: {id}", ss_name="click_element_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def click_toolbar_button(self, table_id: str, button_id: str) -> None:
        """
        Click a button that is part a table specific toolbar.

        Arguments:
            table_id {str} -- Id of the table of the toolbar
            button_id {str} -- Id of the button to be clicked
        """
        self.new_step(action="click_toolbar_button", id=button_id, table_id=table_id)
        if self.is_element(table_id):
            try:
                self.current_element.pressToolbarButton(button_id)
                self.step_pass(
                    msg=f"Successfully clicked toolbar button: {button_id} \
                        for table: {self.current_element.Id}", 
                    ss_name="click_toolbar_button_pass")
            except AttributeError:
                self.current_element.pressButton(button_id)
                self.step_pass(
                    msg=f"Successfully clicked toolbar button: {button_id} for table: {self.current_element.Id}", 
                    ss_name="click_toolbar_button_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while clicking toolbar button: {button_id} for table: {self.current_element.Id}",
                    ss_name="click_toolbar_button_exception",
                    error=err)
                
    @explicit_wait_after(wait_time=__explicit_wait__)
    def double_click(self, id: str, item_id: str, column_id: str) -> None:
        """
        Double click an element that is a child of a SAP GUI shell element.

        Arguments:
            id {str} -- Id of the SAP GUI shell
            item_id {str} -- Id of the element to be double clicked
            column_id {str} -- Id of the parent shell's column where the element to be double clicked is located
        """
        self.new_step(action="double_click", id=id, item_id=item_id, column_id=column_id)
        if self.is_element(id):
            try:
                if self.current_element.Type == "GuiShell":
                    self.current_element.doubleClickItem(item_id, column_id)
                self.step_pass(
                    msg=f"Successfully double clicked id: {self.current_element.Id} at item: {item_id} and column: {column_id}", 
                    ss_name="double_click_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while double clicking: {self.current_element.Id} item: {item_id} and column: {column_id}",
                    ss_name="double_click_exception",
                    error=err)

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_cell_value(self, table_id: str, row_num: int, column_id: str) -> str|None:
        """
        Returns the value of the cell specified by the table, row, and column ids.

        Arguments:
            table_id {str} -- Id of the SAP GUI table element containing the cell
            row_num {int} -- Row id of the cell
            column_id {str} -- Column id of the cell

        Returns:
            str|None -- Returns a string of the cell value or None if the cell is not found
        """
        self.new_step(action="get_cell_value", id=table_id, row_num=row_num, column_id=column_id)
        __value: str = None
        if self.is_element(table_id):
            try:
                __value = self.session.findById(self.current_element.Id).getCellValue(row_num, column_id)
                self.step_pass(msg=f"Success getting cell value from table: {self.current_element.Id} in \
                    column: {column_id} and row: {row_num}", 
                    ss_name="get_cell_value_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception getting cell value from table: {self.current_element.Id} \
                        in column: {column_id} and row: {row_num}",
                    ss_name="get_cell_value_exception",
                    error=err)
        return __value

    @explicit_wait_after(wait_time=__explicit_wait__)
    def set_combobox(self, id: str, key: str) -> None:
        """
        Sets the selection of a SAP GUI combobox element.

        Arguments:
            id {str} -- Id of the combobox element
            key {str} -- Key of the selection to be set
        """
        self.new_step(action="set_combobox", id=id, key=key)
        if self.is_element(id):
            try:
                if self.current_element.Id == "GuiComboBox":
                    self.session.findById(self.current_element.Id).key = key
                    self.step_pass(msg=f"Successfully set combobox: {self.current_element.Id} with key: {key}", 
                        ss_name="set_combobox_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception setting combobox: {self.current_element.Id} with key: {key}",
                    ss_name="set_combobox_exception",
                    error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def get_row_count(self, table_id: str) -> int|None:
        """
        Returns the count of rows of the SAP GUI table specified by the table_id.

        Arguments:
            table_id {str} -- Id of the table on which to count rows

        Returns:
            int|None -- Returns an integer of the number of rows in the table or None if the table is not found
        """
        self.new_step(action="get_row_count", id=table_id)
        __count: int = None
        if self.is_element(table_id):
            try:
                __count = self.current_element.rowCount
                self.step_pass(msg=f"Successfully got count: {__count} from table: {self.current_element.Id}", ss_name="get_row_count_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception getting count from table: {self.current_element.Id}",
                    ss_name="get_row_count_exception",
                    error=err)
        return __count

    @explicit_wait_after(wait_time=__explicit_wait__)
    def get_window_title(self) -> str|None:
        """
        Returns the string value of the currently active main window.

        Returns:
            str|None -- Returns the value of the main window or None if there is no currently active main window
        """
        self.new_step(action="get_window_title")
        self.current_element = self.titl
        __title: str = None
        try:
            __title = self.current_element.Text
            self.step_pass(msg=f"Successfully got window title: {__title} from window: {self.current_element.Id}", ss_name="get_window_title_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"Unhandled exception getting window title for window: {self.current_element.Id}",
                ss_name="get_window_title_exception",
                error=err)
        return __title

    @explicit_wait_after(wait_time=__explicit_wait__)
    def get_value(self, id: str) -> str|None:
        """
        Returns the string value of the SAP GUI element specified by the id. 
        Only applies to elements which have a text property.
        Valid element types: 
            GuiCheckBox
            GuiRadioButton
            GuiComboBox
            GuiTextField
            GuiCTextField
            GuiPasswordField
            GuiLabel
            GuiTitlebar
            GuiStatusbar
            GuiButton
            GuiTab
            GuiShell
            GuiStatusPane

        Arguments:
            id {str} -- Id of the element

        Returns:
            str|None -- Return the value of the element or None if the element if not found or if the element is of a type which does not have a text property
        """
        self.new_step(action="get_value", id=id)
        __value: str = None
        if self.is_element(id):
            try:
                if self.current_element.Type in TextElements:
                    __value = self.current_element.Text
                    self.step_pass(
                        msg=f"Successfully got value from: {self.current_element.Id}", 
                        ss_name="get_value_pass")
                elif self.current_element.Type in ("GuiCheckBox", "GuiRadioButton"):
                    __value = self.current_element.Selected
                    self.step_pass(
                        msg=f"Successfully got value from: {self.current_element.Id}", 
                        ss_name="get_value_pass")
                elif self.current_element.Type == "GuiComboBox":
                    __value = str(self.current_element.Text).strip()
                    self.step_pass(
                        msg=f"Successfully got value from: {self.current_element.Id}", 
                        ss_name="get_value_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception getting value from: {self.current_element.Id}",
                    ss_name="get_value_exception",
                    error=err)
        return __value

    @explicit_wait_after(wait_time=__explicit_wait__)
    def set_text(self, id: str, text: str) -> None:
        """
        Set the text property of the SAP GUI element specified by the id.

        Arguments:
            id {str} -- Id of the element
            text {str} -- Value to set the text property
        """
        self.new_step(action="set_text", id=id, text=text)
        if self.is_element(id):
            try:
                if self.current_element.Type in [i.value for i in TextElements]:
                    self.current_element.Text = text
                    self.step_pass(
                        msg=f"Successfully entered: {text} in: {self.current_element.Id}", 
                        ss_name="set_text_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while entering: {text} in: {self.current_element.Id}",
                    ss_name="set_text_exception",
                    error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def set_cell_value(self, table_id: str, row: int, col: str, text: str) -> None:
        """
        Set the value of a table cell specified by the row and column.

        Arguments:
            table_id {str} -- Id of the SAP GUI table element
            row {int} -- Row number of the cell, zero based
            col {str} -- Name of the table column of the cell
            text {str} -- Value to set the cell
        """
        self.new_step(action="set_cell_value", id=table_id, row=row, col=col, text=text)
        if self.is_element(table_id):
            try:
                self.current_element.modifyCell(row, col, text)
                self.step_pass(msg=f"Successfully input {text} into cell ({row}, {col}).", ss_name="set_cell_value_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while entering: {text} into cell ({row}, {col}",
                    ss_name="set_cell_value_exception",
                    error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def set_checkbox(self, id: str, state: bool) -> None:
        """
        Set the state of a SAP GUI checkbox.

        Arguments:
            id {str} -- Id of the checkbox element
            state {bool} -- State to which to set the checkbox. True is checked and False is unchecked
        """
        self.new_step(action="set_checkbox", id=id, state=state)
        if self.is_element(id):
            try:
                if self.current_element.Type == "GuiCheckBox":
                    self.current_element.selected = state
                    self.step_pass(msg=f"", ss_name="set_checkbox_pass")
                else:
                    self.step_fail(msg=f"", ss_name="set_checkbox_fail")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while selecting checkbox: {self.current_element.Id} in: {self.current_element.Id}",
                    ss_name="set_checkbox_exception",
                    error=err)
        else:
            self.handle_unknown_exception(
                msg=f"Unhandled exception while selecting element: {self.current_element.Id}",
                ss_name="set_checkbox_exception",
                error=err)

    # Buttons & Keys
    @explicit_wait_after(wait_time=__explicit_wait__)
    def send_vkey(self, vkey: str) -> None:
        """
        Send a virtual key press to the currently active main window of the SAP GUi session.
        A complete list of supported virtual keys can be found in Flow.Data.VKEYS

        Arguments:
            vkey {str} -- Virtual key to send to the window
        """
        self.new_step(action="send_vkey", vkey=vkey)
        __vkey_id: str = str(vkey)
        if not __vkey_id.isdigit():
            __search_comb: str = __vkey_id.upper()
            __search_comb = __search_comb.replace(" ", "")
            __search_comb = __search_comb.replace("CONTROL", "CTRL")
            __search_comb = __search_comb.replace("DELETE", "DEL")
            __search_comb = __search_comb.replace("INSERT", "INS")
            try:
                __vkey_id = VKEYS.index(__search_comb)
            except ValueError:
                if __search_comb == "CTRL+S":
                    __vkey_id = 11
                elif __search_comb == "ESC":
                    __vkey_id = 12
                else:
                    self.step_fail(
                        msg=f"Invalid vkey: {__vkey_id}, provide a valid Vkey number or combination", 
                        ss_name="send_vkey_fail")
        try:
            self.main_window.sendVKey(__vkey_id)
            self.step_pass(
                msg=f"Successfully sent vkey: {__vkey_id} to window: {self.main_window.Id}", 
                ss_name="send_vkey_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"Unhandled exception sending vkey: {__vkey_id} to window: {self.main_window.Id}",
                ss_name="send_vkey_exception",
                error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def enter(self) -> None:
        """
        Send a virtual enter key press to the currently active main window of the session.
        """
        self.new_step(action="enter")
        try:
            self.send_vkey(vkey="ENTER")
            self.step_pass(msg=f"Successfully sent ENTER.", ss_name="enter_pass")
        except Exception as err:
            self.handle_unknown_exception(
                    msg=f"Unhandled exception sending ENTER.",
                    ss_name="enter_exception",
                    error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def save(self) -> None:
        """
        Send a virtual Ctrl+S key press sequence to the currently active main window of the session.
        """
        self.new_step(action="save")
        try:
            self.send_vkey(vkey="CTRL+S")
            self.step_pass(msg=f"Successfully sent SAVE.", ss_name="save_pass")
        except Exception as err:
            self.handle_unknown_exception(
                    msg=f"Unhandled exception sending SAVE.",
                    ss_name="save_exception",
                    error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def back(self) -> None:
        """
        Send a virtual F3 key press to the currently active main window of the session.
        """
        self.new_step(action="back")
        try:
            self.send_vkey(vkey="F3")
            self.step_pass(msg=f"Successfully sent BACK.", ss_name="back_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending BACK.", ss_name="back_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f8(self) -> None:
        """
        Send a virtual F8 key press to the currently active main window of the session.
        """
        self.new_step(action="f8")
        try:
            self.send_vkey(vkey="F8")
            self.step_pass(msg=f"Successfully sent F8.", ss_name="f8_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending F8.", ss_name="f8_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f5(self) -> None:
        """
        Send a virtual F5 key press to the currently active main window of the session.
        """
        self.new_step(action="f5")
        try:
            self.send_vkey(vkey="F5")
            self.step_pass(msg=f"Successfully sent F5.", ss_name="f5_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending F5.", ss_name="f5_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f6(self) -> None:
        """
        Send a virtual F6 key press to the currently active main window of the session.
        """
        self.new_step(action="f6")
        try:
            self.send_vkey(vkey="F6")
            self.step_pass(msg=f"Successfully sent F6.", ss_name="f6_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending F6.", ss_name="f6_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f7(self) -> None:
        """
        Send a virtual F7 key press to the currently active main window of the session.
        """
        self.new_step(action="f7")
        try:
            self.send_vkey(vkey="F7")
            self.step_pass(msg=f"Successfully sent F7.", ss_name="f7_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending F7.", ss_name="f7_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f4(self) -> None:
        """
        Send a virtual F4 key press to the currently active main window of the session.
        """
        self.new_step(action="f4")
        try:
            self.send_vkey(vkey="F4")
            self.step_pass(msg=f"Successfully sent F4.", ss_name="f4_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending F4.", ss_name="f4_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f3(self) -> None:
        """
        Send a virtual F3 key press to the currently active main window of the session.
        """
        self.new_step(action="f3")
        try:
            self.send_vkey(vkey="F3")
            self.step_pass(msg=f"Successfully sent F3.", ss_name="f3_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending F3.", ss_name="f3_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f2(self) -> None:
        """
        Send a virtual F2 key press to the currently active main window of the session.
        """
        self.new_step(action="f2")
        try:
            self.send_vkey(vkey="F2")
            self.step_pass(msg=f"Successfully sent F2.", ss_name="f2_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending F2.", ss_name="f2_exception", error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def f1(self) -> None:
        """
        Send a virtual F1 key press to the currently active main window of the session.
        """
        self.new_step(action="f1")
        try:
            self.send_vkey(vkey="F1")
            self.step_pass(msg=f"Successfully sent F1.", ss_name="f1_pass")
        except Exception as err:
            self.handle_unknown_exception(msg=f"Unhandled exception sending F1.", ss_name="f1_exception", error=err)

    # Assertions
    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_element_value_equal(self, id: str, expected_value: str) -> None:
        if self.is_element(id):
            try:
                if self.get_value(id=self.current_element.Id) == expected_value:
                    self.step_pass(
                        msg=f"Assertion equal passed for element: {self.current_element.Id} with \
                            actual value: {self.get_value(id=self.current_element.Id)} \
                            and expected value: {expected_value}", 
                        ss_name="assert_element_value_equal_pass")
                else:
                    self.step_fail(
                        msg=f"Assertion equal failed for element: {self.current_element.Id} with \
                            actual value: {self.get_value(id=self.current_element.Id)} \
                            and expected value: {expected_value}", 
                        ss_name="assert_element_value_equal_fail")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while asserting element: {self.current_element.Id} \
                        equals: {expected_value}",
                    ss_name="assert_element_value_equal_exception",
                    error=err)
        else:
            self.step_fail(
                msg=f"Assertion equal failed, element: {self.current_element.Id} is not present.", 
                ss_name="assert_element_value_equal_fail")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_element_value_not_equal(self, id: str, expected_value: str) -> None:
        if self.is_element(id):
            try:
                if self.get_value(id=self.current_element.Id) != expected_value:
                    self.step_pass(
                        msg=f"Assertion not equal passed for element: {self.current_element.Id} with \
                            actual value: {self.get_value(id=self.current_element.Id)} \
                            and expected value: {expected_value}", 
                        ss_name="assert_element_value_not_equal_pass")
                else:
                    self.step_fail(
                        msg=f"Assertion not equal failed for element: {self.current_element.Id} with \
                            actual value: {self.get_value(id=self.current_element.Id)} \
                            and expected value: {expected_value}", 
                        ss_name="assert_element_value_not_equal_fail")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while asserting element: {self.current_element.Id} \
                        not equal: {expected_value}",
                    ss_name="assert_element_value_not_equal_exception",
                    error=err)
        else:
            self.step_fail(
                msg=f"Assertion not equal failed, element: {self.current_element.Id} is not present.", 
                ss_name="assert_element_value_not_equal_fail")
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def assert_element_present(self, id: str) -> None:
        if self.is_element(id):
            self.step_pass(
                msg=f"Assertion passed, element: {self.current_element.Id} is present.", 
                ss_name="assert_element_present_pass")
        else:
            self.step_fail(
                msg=f"Assertion failed, element: {self.current_element.Id} is not present.", 
                ss_name="assert_element_present_fail")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_element_changeable(self, id: str, expected: bool) -> None:
        if self.is_element(id):
            try:
                if self.current_element.Changeable == expected:
                    self.step_pass(
                        msg=f"Assertion changeable passed for element: {self.current_element.Id}", 
                        ss_name="assert_element_changeable_pass")
                else:
                    self.step_fail(
                        msg=f"Assertion changeable failed for element: {self.current_element.Id}", 
                        ss_name="assert_element_changeable_fail")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while asserting changeability \
                        of element: {self.current_element.Id}",
                    ss_name="assert_element_changeable_exception",
                    error=err)
        else:
            self.step_fail(
                msg=f"Assertion changeable failed, element: {self.current_element.Id} is not present.", 
                ss_name="assert_element_changeable_fail")

    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_element_value_contains(self, id: str, contains_value: str) -> None:
        if self.is_element(id):
            try:
                if contains_value in self.get_value(id=self.current_element.Id):
                    self.step_pass(
                        msg=f"Assertion value contains for element: {self.current_element.Id} with \
                            actual value: {self.get_value(id=self.current_element.Id)} \
                            & expected contains value: {contains_value}", 
                        ss_name="assert_element_value_contains_pass")
                else:
                    self.step_fail(
                        msg=f"Assertion value contains failed for element: {self.current_element.Id} with \
                            actual value: {self.get_value(id=self.current_element.Id)} \
                            & expected value: {contains_value}", 
                        ss_name="assert_element_value_contains_fail")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while asserting element: {self.current_element.Id} \
                        contains: {contains_value}",
                    ss_name="assert_element_value_contains_exception",
                    error=err)
        else:
            self.step_fail(
                msg=f"Assertion value contains failed, element: {self.current_element.Id} is not present.", 
                ss_name="assert_element_value_contains_fail")
    
    @explicit_wait_after(wait_time=__explicit_wait__)
    def assert_success_status(self) -> None:
        try:
            if self.sbar.MessageType == "S":
                self.step_pass(
                    msg=f"Status is success", 
                    ss_name="assert_success_status_pass")
            else:
                self.step_fail(
                    msg=f"Status is {self.sbar.MessageType} -- {self.sbar.Text}", 
                    ss_name="assert_success_status_fail")
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"Unhandled exception while asserting success status",
                ss_name="assert_success_status_exception",
                error=err)

    # Screen Parsing & Visualization
    def visualize_element(self, id: str, visualize: Optional[bool] = False) -> None:
        if self.is_element(id):
            try:
                self.current_element.visualize(visualize) 
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception visualizing element: {self.current_element.Id}",
                    ss_name="visualize_element_exception",
                    error=err)

    # Compound functions
    ## Tables
    def dump_table_values(self, table_id: str) -> Table:
        __table = self.session.FindById(table_id)
        if __table.Type == "GuiTableControl":
            my_table = Table(
                Id = table_id, 
                Type = __table.Type,
                TableObject = __table,
                RowCount = __table.RowCount,
                VisibleRows = __table.VisibleRowCount,
                Columns = [x for x in __table.Columns],
                Rows = [x for x in __table.Rows],
                Data = []
            )
            __columns = [x for x in __table.Columns]
            __rows = [x for x in __table.Rows]
            for row in __rows:
                cells = {}
                for cell in range(0, row.Count):
                    cells[__columns[cell].Name] = row.ElementAt(cell).Text
                my_table.Data.append(cells)
            return my_table
        elif __table.Type == "GuiShell":
            if __table.SubType == "GridView":
                __column_order = __table.ColumnOrder
                my_table = Table(
                    Id = table_id, 
                    Type = __table.SubType,
                    TableObject = __table,
                    RowCount = __table.RowCount,
                    VisibleRows = __table.VisibleRowCount,
                    Columns = __column_order,
                    Rows = [],
                    Data = []
                )
                for row in range(0, __table.RowCount):
                    cells = {}
                    for cell in range(0, __table.ColumnCount):
                        cells[__column_order[cell]] = __table.GetCellValue(row, __column_order[cell])
                    my_table.Data.append(cells)
                return my_table
    
    # Get Table Data
    def get_table_data(self, statement: str) -> Table:
        fields, max_rows, table, conditions = self.select_parse(statement)
        self.start_transaction(transaction="SE16")
        
        # Set table
        self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtDATABROWSE-TABLENAME", text=table)
        self.enter()
        
        # Set conditions
        self.click_element(id="/app/con[0]/ses[0]/wnd[0]/mbar/menu[3]/menu[2]")
        self.click_element(id="/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[14]")  # Unselect All
        for condition in conditions:
            self.click_element(id="/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[71]")  # Search
            self.set_text(id="/app/con[0]/ses[0]/wnd[2]/usr/txtRSYSF-STRING", text=condition[0])
            self.set_checkbox(id="/app/con[0]/ses[0]/wnd[2]/usr/chkSCAN_STRING-START", state=False)
            self.click_element(id="/app/con[0]/ses[0]/wnd[2]/tbar[0]/btn[0]")
            self.session.FindById("/app/con[0]/ses[0]/wnd[3]/usr/lbl[3,2]").SetFocus()
            self.click_element(id="/app/con[0]/ses[0]/wnd[3]/tbar[0]/btn[2]")
            self.set_checkbox(id="/app/con[0]/ses[0]/wnd[1]/usr/chk[2,6]", state=True)
            self.click_element(id="/app/con[0]/ses[0]/wnd[1]/usr/chk[2,6]")
            self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtI1-LOW", text=condition[2])
            self.f2()
            gv = self.session.FindById("/app/con[0]/ses[0]/wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell")
            # Set selection option
        
        # Set max rows to return
        self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/txtMAX_SEL", text=max_rows)
        
        # Set fields
        self.click_element(id="/app/con[0]/ses[0]/wnd[0]/mbar/menu[3]/menu[0]/menu[1]")
        self.click_element(id="/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[14]")
        for field in fields:
            self.click_element(id="/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[71]")
            self.set_text(id="/app/con[0]/ses[0]/wnd[2]/usr/txtRSYSF-STRING", text="HERKL")
            self.set_checkbox(id="/app/con[0]/ses[0]/wnd[2]/usr/chkSCAN_STRING-START", state=False)
            self.click_element(id="/app/con[0]/ses[0]/wnd[2]/tbar[0]/btn[0]")
            self.session.FindById("/app/con[0]/ses[0]/wnd[3]/usr/lbl[3,2]").SetFocus()
            self.click_element(id="/app/con[0]/ses[0]/wnd[3]/tbar[0]/btn[2]")
            self.set_checkbox(id="/app/con[0]/ses[0]/wnd[1]/usr/chk[1,3]", state=True)
            self.click_element(id="/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[6]")
    
    ## Sales Orders
    def availability_control(self) -> None:
        if self.is_element("usr/btnBUT3"):
            try:
                if "availability" in self.titl.Text.lower():
                    try:
                        self.click_element(id=self.current_element.Id)
                    except Exception as err2:
                        self.logger.log.debug(f"Availability control error|{err2}")
            except Exception as err:
                self.handle_unknown_exception(
                    f"Unable to process availability control", 
                    ss_name="availability_control_exception", 
                    error=err)
    
    def fill_va01_initial_screen(self, order_type: str, sales_org: str, dist_ch: str, division: str, sales_office: Optional[str] = "", sales_group: Optional[str] = "") -> None:
        self.set_text(id="usr/ctxtVBAK-AUART", text=order_type)
        self.set_text(id="usr/ctxtVBAK-VKORG", text=sales_org)
        self.set_text(id="usr/ctxtVBAK-VTWEG", text=dist_ch)
        self.set_text(id="usr/ctxtVBAK-SPART", text=division)
        self.set_text(id="usr/ctxtVBAK-VKBUR", text=sales_office)
        self.set_text(id="usr/ctxtVBAK-VKGRP", text=sales_group)
        self.enter()

    def fill_va01_header(self, sold_to: str, ship_to: str, customer_reference: Optional[str] = None, customer_reference_date: Optional[str] = None) -> None:
        self.set_text(id="usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR", text=sold_to)
        self.set_text(id="usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR", text=ship_to)
        if customer_reference is not None:
            self.set_text(id="usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=customer_reference)
        else:
            self.set_text(id="usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", text=f"PO_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")
        if customer_reference_date is not None:
            self.set_text(id="usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK", text=customer_reference_date)
        self.enter()

    def fill_va01_line_items(self, line_items: list[dict]) -> None:
        """
        Fills a sales order in transaction VA01 with line item data provided as a list of key value pairs from the line_items attribute.

        Arguments:
            line_items {list[dict]} -- A list of key-value pairs of line item data like:
                "items": [
                    {
                        "material": 1223344556,
                        "qty": 10,
                        "uom": "PC", 
                        "customer_material": "123456789",
                        "item_category": "TAN"
                        "amount": "10.00", 
                        "storage_location": "1234",
                        "shipping_point": "6789",
                    }
                ]
        """
        self.click_element(id="usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01")
        for item in line_items:
            self.set_text(id="usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]", text=item.get('material'))
            self.set_text(id="usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]", text=item.get('qty'))
            self.set_text(id="usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-VRKME[3,0]", text=item.get('uom'))
            if "net_price" in item.keys():
                self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[23,0]", text=item.get('amount'))
            if "customer_material" in item.keys():
                self.set_text(id="usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-KDMAT[12,0]", text=item.get('customer_material'))
            if "item_category" in item.keys():
                self.set_text(id="usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PSTYV[7,0]", text=item.get('item_category'))
            if "shipping_point" in item.keys():
                self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-VSTEL[70,0]", text=item.get('shipping_point'))
            if "storage_location" in item.keys():
                self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[67,0]", text=item.get('storage_location'))
            self.enter()
    
    ## Delivery
    def display_delivery(self, delivery: str) -> None:
        self.start_transaction("VL03N")
        self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtLIKP-VBELN", text=delivery)
        self.wait(0.5)
        self.enter()

    def get_delivery_header_outputs(self, delivery: str) -> list:
        self.display_delivery(delivery=delivery)
        self.click_element(id="/app/con[0]/ses[0]/wnd[0]/mbar/menu[3]/menu[2]/menu[0]")

    def fill_vl01n_initial_screen(self, shipping_point: str, sales_order: str, selection_date: Optional[str] = None) -> None:
        self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtLIKP-VSTEL", text=shipping_point)
        self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtLV50C-VBELN", text=sales_order)
        if selection_date is not None:
            self.set_text(id="/app/con[0]/ses[0]/wnd[0]/usr/ctxtLV50C-DATBI", text=selection_date)
        self.enter()
    
    ## Selenium Web based functions
    web_keys = Keys
    
    def web_session(
        self,
        browser: Optional[BrowserType] = BrowserType.CHROME, 
        headless: Optional[bool] = False, 
        insecure_certs: Optional[bool] = True, 
        log_level: Optional[int] = 3,
        load_strategy: Optional[str] = "normal"
        ) -> None:
        options = None
        if browser == BrowserType.CHROME:
            options = webdriver.ChromeOptions()
            options.page_load_strategy = load_strategy
            options.acceptInsecureCerts = insecure_certs
            if headless:
                options.add_argument("--headless")
            options.add_argument(f"--log-level={log_level}")
            self.web_driver = webdriver.Chrome(options=options)
        else:
            options = webdriver.ChromeOptions()
            options.page_load_strategy = load_strategy
            options.acceptInsecureCerts = insecure_certs
            if headless:
                options.add_argument("--headless")
            options.add_argument(f"--log-level={log_level}")
            self.web_driver = webdriver.Chrome(options=options)
        self.web_main_window_handle = self.web_driver.current_window_handle
        self.web_driver.maximize_window()
    
    def web_find_by_xpath(self, xpath: str, return_element: bool = False, wait_time: Optional[float] = None) -> Any:
        __wait_time = wait_time if wait_time is not None else self.web_wait
        self.web_element = None
        try:
            self.web_element = WebDriverWait(self.web_driver, __wait_time).until(lambda x: x.find_element(by=By.XPATH, value=xpath))
        except Exception as e:
            self.documentation(f"UNHANDLED ERROR: {e}")
        if return_element:
            return self.web_element
    
    def web_get_value(self, xpath: str, wait_time: Optional[float] = None) -> str:
        __text = None
        try:
            self.web_find_by_xpath(xpath=xpath, wait_time=wait_time)
            try:
                __text = self.web_element.text
            except:
                __text = self.web_element.get_attribute('value')
        except:
            self.documentation(f"Unable to get text from web element: {xpath}")
        return __text
    
    def web_click_element(self, xpath: str, wait_time: Optional[float] = None) -> None:
        self.web_find_by_xpath(xpath=xpath, wait_time=wait_time)
        self.web_driver.execute_script("arguments[0].click();", self.web_element)
    
    def web_wait_for_element(self, xpath: str, timeout: Optional[float] = 5.0, delay_time: Optional[float] = 1.0, wait_time: Optional[float] = None) -> None:
        t = Timer()
        while t.elapsed() <= timeout:
            self.web_find_by_xpath(xpath=xpath, wait_time=wait_time)
            try:
                if not self.web_element.is_displayed():
                    self.wait(delay_time)
                else:
                    break
            except Exception as err:
                continue
    
    def web_set_text(self, xpath: str, text: str) -> None:
        self.web_find_by_xpath(xpath=xpath)
        self.web_element.clear()
        self.web_element.send_keys(text)
    
    def web_set_iframe_active(self, xpath: str) -> None:
        self.iframe = None
        self.iframe = self.web_find_by_xpath(xpath=xpath, return_element=True)
        if self.iframe is not None:
            try:
                self.web_driver.switch_to.frame(self.iframe)
            except:
                # Switch back to parent frame in case of error during child frame action
                self.web_driver.switch_to.parent_frame()
    
    def web_set_iframe_inactive(self) -> None:
        self.web_driver.switch_to.parent_frame()
        self.web_iframe = None
    
    def web_set_zoom(self, zoom: int|float) -> None:
        self.web_driver.execute_script(f"document.body.style.zoom='{zoom}%'")
    
    def web_open_url(self, url: str) -> None:
        self.web_driver.get(url)
    
    def web_enter(self, xpath: Optional[str] = None) -> None:
        if xpath is not None:
            self.web_find_by_xpath(xpath=xpath)
        self.web_element.send_keys(Keys.ENTER)
    
    def web_exit(self) -> None:
        self.web_driver.close()
