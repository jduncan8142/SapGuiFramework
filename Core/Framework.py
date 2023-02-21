from typing import Any, Optional
import win32com.client
from CRUD.Interface import DB
from Flow.Data import Case, TextElements, VKEYS
from Flow.Results import Result
from Flow.Actions import Step
from Logging.Logging import Logger, LoggingConfig
from Core.Utilities import *
from Core.SAP import *
from time import sleep
import atexit
import base64


class Session:
    __explicit_wait__: float = 0.0
    
    def __init__(self) -> None:
        self.case: Case = Case()
        self.logger: Logger = None
        if self.case.LogConfig is None:
            self.logger = Logger(config=DB().db["LoggingConfig"])
        else:
            self.logger = Logger(config=self.case.LogConfig)
        Session.__explicit_wait__ = self.case.ExplicitWait
        self.__connection_number: int = 0
        self.__session_number: int = 0
        self.__window_number: int = 0
        self.connection_name: str = None
        self.sap_gui: win32com.client.CDispatch = None
        self.sap_app: win32com.client.CDispatch = None
        self.connection: win32com.client.CDispatch = None
        self.session: win32com.client.CDispatch = None
        self.session_info: win32com.client.CDispatch = None
        self.main_window: win32com.client.CDispatch = None
        self.mbar: win32com.client.CDispatch = None
        self.tbar0: win32com.client.CDispatch = None
        self.titl: win32com.client.CDispatch = None
        self.tbar1: win32com.client.CDispatch = None
        self.usr: win32com.client.CDispatch = None
        self.sbar: win32com.client.CDispatch = None
        self.current_element: win32com.client.CDispatch = None
        self.current_transaction: str = None
        self.current_step: Step = None
        atexit.register(self.cleanup)
    
    def __post_init__(self) -> None:
        if not self.current_step:
            Step(
                Action="Create Session", 
                ElementId="", 
                Args=[],
                Name="Create New Session", 
                Description="Creates and return a new SAP session object.")
    
    def exit(self) -> None:
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
        if self.case.CloseSAPOnCleanup:
            self.exit()
        if self.case.Status.Result is None:
            if len(self.case.Status.FailedSteps) != 0:
                self.case.Status.Result = Result.FAIL
        else:
            self.case.Status.Result = Result.PASS
        self.documentation(
            f"{self.case.Name} completed with \
                status: {self.case.Status.Result.value}")
    
    def hard_copy(
        self, 
        filename: str, 
        image_type: Optional[str] = "PNG", 
        pos: Optional[tuple[int, int, int, int]] = None
    ) -> bytes:
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
                    return base64.b64encode(f_img.read())
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during hard_copy", 
                ss_name="hard_copy_exception", 
                error=err)

    @explicit_wait_before(wait_time=__explicit_wait__)
    def capture_fullscreen(self, screenshot_name: str) -> bytes:
        shot_bytes: bytes = None
        try:
            shot_bytes = self.hard_copy(screenshot_name, "PNG")
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during screen capture", 
                ss_name="take_screenshot_exception", 
                error=err)
        return shot_bytes
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def capture_region(
        self, 
        screenshot_name: str, 
        pos: tuple[int, int, int, int]
    ) -> bytes:
        shot_bytes: bytes = None
        try:
            shot_bytes = self.hard_copy(screenshot_name, "PNG", pos)
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during screen capture", 
                ss_name="take_screenshot_exception", 
                error=err)
        return shot_bytes
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def capture_element(
        self, 
        screenshot_name: str, 
        element_id: str
    ) -> bytes:
        shot_bytes: bytes = None
        try:
            __element = self.session.FindById(self.ace_id(element_id))
            __pos = (
                __element.ScreenLeft, 
                __element.ScreenTop, 
                __element.Width, 
                __element.Height
            )
            shot_bytes = self.hard_copy(__element.Name, "PNG", __pos)
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during screen capture", 
                ss_name="take_screenshot_exception", 
                error=err)
        return shot_bytes

    def ace_id(self, id: Optional[str] = None) -> str:
        base_id: str = f"/app/con[{self.__connection_number}]/ses\
            [{self.__session_number}]/wnd[{self.__window_number}]"
        if id in ("",  " ", None):
            return base_id
        elif id.startswith("usr"):
            return f"{base_id}/{id}"
        elif id.startswith("/usr"):
            return f"{base_id}{id}"
        elif id.startswith("wnd"):
            return f"/app/con[{self.__connection_number}]/ses\
                [{self.__session_number}]/{id}"
        elif id.startswith("/wnd"):
            return f"/app/con[{self.__connection_number}]/ses\
                [{self.__session_number}]{id}"
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
        _msg = msg if msg is not None else f"{self.current_step.Name} \
            -- {self.current_step.Description}"
        if _msg is not None and _msg != "" and _msg != "--":
            self.logger.log.documentation(_msg)
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def step_fail(
        self, 
        msg: Optional[str] = None, 
        ss_name: Optional[str] = None, 
        error: Optional[str] = None
    ) -> None:
        if msg:
            self.logger.log.error(msg)
        self.current_step.Status.Result = Result.FAIL
        self.current_step.Status.Error = error if error is not None else ""
        self.case.Status.FailedSteps.append(self.current_step)
        if self.case.ScreenShotOnFail:
            __ss_name = ss_name if ss_name is not None else f"screenshot\
                _{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
            self.case.Status.FailedScreenShots.append(
                self.capture_fullscreen(screenshot_name=__ss_name)
            )
        if self.case.ExitOnFail:
            sys.exit()

    @explicit_wait_before(wait_time=__explicit_wait__)
    def step_pass(
        self, 
        msg: Optional[str] = None, 
        ss_name: Optional[str] = None
    ) -> None:
        if msg:
            self.logger.log.info(msg)
        self.current_step.Status.Result = Result.PASS
        self.case.Status.PassedSteps.append(self.current_step)
        if self.case.ScreenShotOnPass:
            __ss_name = ss_name if ss_name is not None else f"screenshot\
                _{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
            self.case.Status.PassedScreenShots.append(
                self.capture_fullscreen(screenshot_name=__ss_name)
            )
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def handle_unknown_exception(
        self, 
        msg: Optional[str] = None, 
        ss_name: Optional[str] = None, 
        error: Optional[str] = None
    ) -> None:
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
        data: Optional[dict] = None
    ) -> None:
        __name = name if name is not None else Case.default_name
        __desc = desc if desc is not None else Case.empty_string
        __bpo = bpo if bpo is not None else Case.default_business_process_owner
        __ito = ito if ito is not None else Case.default_it_owner
        __doc_link = doc_link if doc_link is not None else Case.empty_string
        __case_path = case_path if case_path is not None else Case.default_case_path
        __log_config = log_config if log_config is not None else Case.default_log_config
        __date_format = date_format if date_format is not None else Case.default_date_format
        __explicit_wait = explicit_wait if explicit_wait is not None else Case.default_explicit_wait
        __screenshot_on_pass = screenshot_on_pass if screenshot_on_pass is not None else Case.ScreenShotOnPass
        __screenshot_on_fail = screenshot_on_fail if screenshot_on_fail is not None else Case.ScreenShotOnFail
        __fail_on_error = fail_on_error if fail_on_error is not None else Case.FailOnError
        __exit_on_fail = exit_on_fail if exit_on_fail is not None else Case.ExitOnFail
        __close_on_cleanup = close_on_cleanup if close_on_cleanup is not None else Case.CloseSAPOnCleanup
        __system = system if system is not None else Case.default_system
        __steps = steps if steps is not None else Case.empty_list_factory
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
    
    def new_step(
        self, 
        action: str, 
        id: Optional[str] = "", 
        name: Optional[str] = None, 
        desc: Optional[str] = None, 
        *args
    ) -> None:
        __action = action.title()
        __name = name if name is not None else __action
        __desc = desc if desc is not None else ""
        self.current_step = Step(
            Action = __action, 
            ElementId = id, 
            Args = args, 
            Name = __name, 
            Description = __desc)
        self.collect_step_meta_data()
        self.case.Steps.append(self.current_step)
    
    def open_connection(self, connection_name: str) -> None:
        if not self.current_step:
            self.new_step(
                action="Open Connection", 
                name="Open New Connection", 
                desc="Opens a new SAP scripting connection to the provided instance name.")
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
                self.connection = self.sap_app.OpenConnection(self.connection_name, True)
                self.session = self.connection.children(self.__session_number)
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

    def connect_existing_session(self) -> None:
        #TODO: Port connect_to_session & connect_to_existing_connection together
        pass    

    @explicit_wait_before(wait_time=__explicit_wait__)
    def collect_session_info(self) -> None:
        try:
            if self.session:
                self.wait_for_element(self.ace_id())
                self.main_window = self.session.findById(self.ace_id())
                self.mbar = self.session.findById(f"{self.ace_id()}/mbar")
                self.tbar0 = self.session.findById(f"{self.ace_id()}/tbar[0]")
                self.titl = self.session.findById(f"{self.ace_id()}/titl")
                self.tbar1 = self.session.findById(f"{self.ace_id()}/tbar[1]")
                self.usr = self.session.findById(f"{self.ace_id()}/usr")
                self.sbar = self.session.findById(f"{self.ace_id()}/sbar")
                self.session_info = self.session.info
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"Unhandled exception while collecting session info|{err}", 
                ss_name="collect_session_info_exception", 
                error=err)
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def collect_step_meta_data(self) -> None:
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
            self.handle_unknown_exception(
                msg=f"Unhandled exception while collecting step metadata|{err}", 
                ss_name="collect_step_metadata_exception", 
                error=err)
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def collect_case_meta_data(self) -> None:
        try:
            if self.case and self.session:
                self.case.SapMajorVersion = self.sap_app.MajorVersion
                self.case.SapMinorVersion = self.sap_app.MinorVersion
                self.case.SapPatchLevel = self.sap_app.PatchLevel
                self.case.SapRevision = self.sap_app.Revision
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"Unhandled exception while collecting case metadata|{err}", 
                ss_name="collect_case_metadata_exception", 
                error=err)
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def collect_sbar_element(self) -> None:
        try:
            self.current_step.StatusBar = GuiStatusbar(
                Instance=self.sbar,
                Id=self.sbar.Id,
                Name=self.sbar.Name,
                Text=self.sbar.Text,
                ScreenLeft=self.sbar.ScreenLeft,
                ScreenTop=self.sbar.ScreenTop,
                Handle=self.sbar.Handle,
                Left=self.sbar.Left,
                Top=self.sbar.Top,
                Height=self.sbar.Height,
                Width=self.sbar.Width,
                Tooltip=self.sbar.Tooltip,
                DefaultTooltip=self.sbar.DefaultTooltip,
                IconName=self.sbar.IconName,
                Key=self.sbar.Key,
                Changeable=self.sbar.Changeable,
                ContainerType=self.sbar.ContainerType,
                MessageId=self.sbar.MessageId,
                MessageNumber=self.sbar.MessageNumber,
                MessageType=self.sbar.MessageType,
                Pane0=self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]"),
                Pane1=self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[1]"),
                Pane2=self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[2]"),
                Pane3=self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[3]"),
                Pane4=self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[4]"),
                Pane5=self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[5]"),
                Pane6=self.session.findById("/app/con[0]/ses[0]/wnd[0]/sbar/pane[6]"))
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"Unhandled exception while collecting statusbar element", 
                ss_name="collect_sbar_element_exception", 
                error=err)
    
    def start_transaction(self, transaction: str) -> None:
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
        try:
            self.session.endTransaction()
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception during end_transaction.", 
                ss_name="end_transaction_exception", 
                error=err)
            
    def is_element(self, element: str) -> bool:
        try:
            __element = self.ace_id(element)
            self.current_element = self.session.findById(__element)
            self.step_pass(
                msg=f"Element: {__element} is valid", 
                ss_name="is_element_pass")
            return True
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"SAP element id not found: {__element}", 
                ss_name="is_element_exception",
                error=err)
        return False
    
    @explicit_wait_before(wait_time=__explicit_wait__)
    def set_v_scrollbar(self, id: str, pos: int) -> None:
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
    def get_v_scrollbar(self, id: str) -> int | None:
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
    def get_h_scrollbar(self, id: str) -> int | None:
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
    
    def wait(self, seconds: float) -> None:
        if seconds == 1.0:
            self.documentation(f"Waiting 1 second...")
        else:
            self.documentation(f"Waiting {seconds} seconds...")
        sleep(secs=seconds)
    
    def wait_for_element(self, id: str, timeout: Optional[float] = 60.0) -> None:
        try:
            __id = self.ace_id(id)
            t = Timer()
            while True:
                if not self.is_element(element=__id) and t.elapsed() <= timeout:
                    self.wait(value=0.5)
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

    @explicit_wait_before(wait_time=__explicit_wait__)
    def maximize_window(self) -> None:
        try:
            self.main_window.maximize()
            self.step_pass(msg="Window maximized.", ss_name="maximize_window_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg="Unhandled exception maximizing window.", 
                ss_name="maximize_window_exception", 
                error=err)

    def restart_session(self, delay: Optional[float] = 0.0) -> None:
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

    @explicit_wait_after(wait_time=__explicit_wait__)
    def click_element(self, id: str) -> None:
        if self.is_element(id):
            try:
                if self.current_element.Type in ("GuiTab", "GuiMenu"):
                    self.current_element.select()
                    self.step_pass(
                        msg=f"Successfully clicked element: {self.current_element.Id}", 
                        ss_name="click_element_success")
                elif self.current_element.Type == "GuiButton":
                    self.current_element.press()
                    self.step_pass(
                        msg=f"Successfully clicking GuiButton: {self.current_element.Id}", 
                        ss_name="click_gui_button_success")
                else:
                    self.step_fail(
                        msg=f"Unable to click element: {self.current_element.Id}", 
                        ss_name="click_element_failed")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while clicking element: {self.current_element.Id}",
                    ss_name="click_element_exception",
                    error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def click_toolbar_button(self, table_id: str, button_id: str) -> None:
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
                    msg=f"Successfully clicked toolbar button: {button_id} \
                        for table: {self.current_element.Id}", 
                    ss_name="click_toolbar_button_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while clicking toolbar button: {button_id} \
                        for table: {self.current_element.Id}",
                    ss_name="click_toolbar_button_exception",
                    error=err)
                
    @explicit_wait_after(wait_time=__explicit_wait__)
    def double_click(self, id: str, item_id: str, column_id: str) -> None:
        if self.is_element(id):
            try:
                if self.current_element.Type == "GuiShell":
                    self.current_element.doubleClickItem(item_id, column_id)
                self.step_pass(
                    msg=f"Successfully double clicked id: {self.current_element.Id} at \
                        item: {item_id} and column: {column_id}", 
                    ss_name="double_click_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception while double clicking: {self.current_element.Id} \
                        item: {item_id} and column: {column_id}",
                    ss_name="double_click_exception",
                    error=err)

    @explicit_wait_before(wait_time=__explicit_wait__)
    def get_cell_value(self, table_id: str, row_num: int, column_id: str) -> str | None:
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
        if self.is_element(id):
            try:
                if self.current_element.Id == "GuiComboBox":
                    self.session.findById(self.current_element.Id).key = key
                    self.step_pass(msg=f"Successfully set combobox: {self.current_element.Id} \
                        with key: {key}", 
                        ss_name="set_combobox_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception setting combobox: {self.current_element.Id} \
                        with key: {key}",
                    ss_name="set_combobox_exception",
                    error=err)

    @explicit_wait_after(wait_time=__explicit_wait__)
    def get_row_count(self, table_id: str) -> int | None:
        __count: int = None
        if self.is_element(table_id):
            try:
                __count = self.current_element.rowCount
                self.step_pass(msg=f"Successfully got count: {__count} from \
                    table: {self.current_element.Id}", ss_name="get_row_count_pass")
            except Exception as err:
                self.handle_unknown_exception(
                    msg=f"Unhandled exception getting count from table: {self.current_element.Id}",
                    ss_name="get_row_count_exception",
                    error=err)
        return __count

    @explicit_wait_after(wait_time=__explicit_wait__)
    def get_window_title(self) -> str | None:
        self.current_element = self.titl
        __title: str = None
        try:
            __title = self.current_element.Text
            self.step_pass(msg=f"Successfully got window title: {__title} from \
                window: {self.current_element.Id}", ss_name="get_window_title_pass")
        except Exception as err:
            self.handle_unknown_exception(
                msg=f"Unhandled exception getting window title for window: {self.current_element.Id}",
                ss_name="get_window_title_exception",
                error=err)
        return __title

    @explicit_wait_after(wait_time=__explicit_wait__)
    def get_value(self, id: str) -> str | None:
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

    def try_and_continue(self, func: str, *args, **kwargs) -> Any:
        __result = None
        try:
            if hasattr(self, func) and callable(func := getattr(self, func)):
                __result = func(*args, **kwargs)
        except Exception as err:
            self.logger.log.info(f"Unhandled exception during Try and Continue \
                wrapped function: {func}")
            self.current_step.Status.Result = Result.WARN
            self.current_step.Status.Error = err
            self.case.Status.PassedSteps.append(self.current_step)
            if self.case.ScreenShotOnPass:
                self.case.Status.PassedScreenShots.append(
                    self.capture_fullscreen(
                        screenshot_name="try_and_continue_exception"
                    )
                )            
        return __result

    @explicit_wait_after(wait_time=__explicit_wait__)
    def send_vkey(self, vkey: str) -> None:
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

    # Buttons & Keys
    @explicit_wait_after(wait_time=__explicit_wait__)
    def enter(self) -> None:
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
        try:
            self.send_vkey(vkey="F3")
            self.step_pass(msg=f"Successfully sent BACK.", ss_name="back_pass")
        except Exception as err:
            self.handle_unknown_exception(
                    msg=f"Unhandled exception sending BACK.",
                    ss_name="back_exception",
                    error=err)
