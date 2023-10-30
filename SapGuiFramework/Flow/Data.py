from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum, auto
from pathlib import Path
from typing import Optional
import xml.etree.ElementTree as ET
from Core.Utilities import get_main_dir
from Flow.Actions import Step
from Flow.Results import ResultCase
from Logging.Logging import LoggingConfig
from dotenv import load_dotenv
import json
from types import SimpleNamespace
import os


class CaseTypes(Enum):
    GUI = auto()
    WEB = auto()


@dataclass
class Condition:
    Type:str
    Value: str
    Currency: str


@dataclass
class SalesOrderHeader:
    OrderType: str
    SalesOrg: str
    DistCh: str
    Division: str
    SoldTo: str
    ShipTo: str
    CustReference: str = None
    CustRefDate: Optional[str] = None
    ShippingCondition: Optional[str] = None
    RequestedDeliveryDate: Optional[str] = None
    CompleteDelivery: Optional[bool] = None
    DeliveryBlock: Optional[str] = None
    BillingBlock: Optional[str] = None
    PaymentTerms: Optional[str] = None
    IncoVersion: Optional[str] = None
    Incoterms: Optional[str] = None
    IncoLocation1: Optional[str] = None
    IncoLocation2: Optional[str] = None
    OrderReason: Optional[str] = None
    DeliveryPlant: Optional[str] = None
    TotalWeight: Optional[str] = None
    Volume: Optional[str] = None
    PricingDate: Optional[str] = None


@dataclass
class SalesOrderItem:
    Material: str
    Qty: str
    Uom: str
    LineNumber: Optional[str] = None
    ItemDescription: Optional[str] = None
    ItemCategory: Optional[str] = None
    WbsElement: Optional[str] = None
    ReasonForRejection: Optional[str] = None
    CustMaterialNum: Optional[str] = None
    CustReference: Optional[str] = None
    POItem: Optional[str] = None
    FirstDate: Optional[str] = None
    Plant: Optional[str] = None
    StorageLocation: Optional[str] = None
    ShippingPoint: Optional[str] = None
    ProfitCenter: Optional[str] = None
    PricingConditions: Optional[list[Condition]] = None


@dataclass
class SalesOrder:
    Header: SalesOrderHeader
    Items: list[SalesOrderItem]


@dataclass
class Systems:
    def get_sap_systems() -> list[str]:
        __tree = ET.parse(Systems.LandscapeXML)
        __root = __tree.getroot()
        for child in __root:
            if child.tag == "Services":
                return [gc.attrib['name'] for gc in child]
    
    LandscapeXML: str
    AvailableSystems: list[str] = field(default_factory=get_sap_systems)


@dataclass
class Case:    
    def default_explicit_wait() -> float:
        return 0.25
    
    def default_date_format() -> str:
        return "%m/%d/%Y"
    
    def default_name() -> str:
        return f"test_{datetime.now().strftime('%m%d%Y_%H%M%S')}"
    
    def default_log_config() -> LoggingConfig:
        return LoggingConfig()
    
    def default_result() -> ResultCase:
        return ResultCase()

    def default_case_type() -> CaseTypes:
        return CaseTypes.GUI
    
    def default_web_wait() -> float:
        return 1.0
    
    Name: str = field(default_factory=default_name)
    Description: str = field(default_factory=str)
    BusinessProcessOwner: str = field(default_factory=str)
    ITOwner: str = field(default_factory=str)
    DocumentationLink: str = field(default_factory=str)
    
    CasePath: Path|str = field(default_factory=get_main_dir)
    LogConfig: LoggingConfig = field(default_factory=default_log_config)
    DateFormat: str = field(default_factory=default_date_format)
    ExplicitWait: float = field(default_factory=default_explicit_wait)
    CaseType: CaseTypes = field(default_factory=default_case_type)
    
    WebWait: float = field(default_factory=default_web_wait)
    
    ScreenShotOnPass: bool = False
    ScreenShotOnFail: bool = False
    FailOnError: bool = True
    ExitOnFail: bool = True
    CloseSAPOnCleanup: bool = True
    
    Systems: dict = field(default_factory=dict)
    Steps: list[Step] = field(default_factory=list)
    Data: dict = field(default_factory=dict)
    Status: ResultCase = field(default_factory=default_result)
    
    SapMajorVersion: Optional[int] = None
    SapMinorVersion: Optional[int] = None
    SapPatchLevel: Optional[int] = None
    SapRevision: Optional[int] = None


def load_case_from_excel_file(excel_file: str|Path) -> Case:
    """
    Load test case data from a excel file.

    Arguments:
        data_file {str|Path} -- Path the excel data file

    Keyword Arguments:
        case {Optional[Case]} -- An existing Case object that will be updated 
                                from the loaded excel file (default: {None})

    Returns:
        Case -- Return the updated Case object.
    """
    raise NotImplementedError


def load_case_from_json_file(data_file: str) -> Case:
    """
    Load test case data from a json file.

    Arguments:
        data_file {str} -- Path the json data file

    Returns:
        Case -- Return a Case object.
    """
    __case = Case()
    #Load test case
    with open(file=data_file, mode="r") as f:
        __case = json.load(fp=f, object_hook=lambda d: SimpleNamespace(**d))
    return __case

    # __data: dict = json.load(open(data_file, "rb"))
    # return load_case(data=__data, case=case)


def load_case(data: dict, case: Case) -> Case:
    """
    Load test case data from dict. If a value does not exist in the dict
    attempt to get it from an environment variable or load the default value.

    Arguments:
        data {dict} -- dict of data values
        case {Case} -- An existing Case object that will be updated 
                        from the loaded data dict

    Returns:
        Case -- Return the updated Case object.
    """
    load_dotenv()
    _case = case if case is not None else Case()
    __data: dict = data
    if "case_name" in __data:
        _case.Name = __data.get("case_name")
    elif "case_name" in os.environ:
        _case.Name = os.getenv(os.getenv("case_name"))
    else:
        _case.Name = f"test_{datetime.datetime.now().strftime('%m%d%Y_%H%M%S')}"
    if "description" in __data:
        _case.Description = __data.get("description")
    elif "description" in os.environ:
        _case.Description = os.getenv(os.getenv("description"))
    else:
        _case.Description = ""
    if "business_owner" in __data:
        _case.BusinessProcessOwner = __data.get("business_owner")
    elif "business_owner" in os.environ:
        _case.BusinessProcessOwner = os.getenv(os.getenv("business_owner"))
    else:
        _case.BusinessProcessOwner = "Business Process Owner"
    if "it_owner" in __data:
        _case.ITOwner = __data.get("it_owner")
    elif "it_owner" in os.environ:
        _case.ITOwner = os.getenv(os.getenv("it_owner"))
    else:
        _case.ITOwner = "Technical Owner"
    if "doc_link" in __data:
        _case.DocumentationLink = __data.get("doc_link")
    elif "doc_link" in os.environ:
        _case.DocumentationLink = os.getenv(os.getenv("doc_link"))
    else:
        _case.DocumentationLink = ""
    if "case_path" in __data:
        _case.CasePath = __data.get("case_path")
    elif "case_path" in os.environ:
        _case.CasePath = os.getenv(os.getenv("case_path"))
    else:
        _case.CasePath = ""
    if "date_format" in __data:
        _case.DateFormat = __data.get("date_format")
    elif "date_format" in os.environ:
        _case.DateFormat = os.getenv(os.getenv("date_format"))
    else:
        _case.DateFormat = "%m/%d/%Y"
    if "explicit_wait" in __data:
        _case.ExplicitWait = __data.get("explicit_wait")
    elif "explicit_wait" in os.environ:
        _case.ExplicitWait = os.getenv(os.getenv("explicit_wait"))
    else:
        _case.ExplicitWait = 0.25
    if "web_wait" in __data:
        _case.WebWait = __data.get("web_wait")
    elif "explicit_wait" in os.environ:
        _case.WebWait = os.getenv(os.getenv("web_wait"))
    else:
        _case.WebWait = 1.0
    if "screenshot_on_pass" in __data:
        _case.ScreenShotOnPass = __data.get("screenshot_on_pass")
    elif "screenshot_on_pass" in os.environ:
        _case.ScreenShotOnPass = os.getenv(os.getenv("screenshot_on_pass"))
    else:
        _case.ScreenShotOnPass = False
    if "screenshot_on_fail" in __data:
        _case.ScreenShotOnFail = __data.get("screenshot_on_fail")
    elif "screenshot_on_fail" in os.environ:
        _case.ScreenShotOnFail = os.getenv(os.getenv("screenshot_on_fail"))
    else:
        _case.ScreenShotOnFail = False
    if "fail_on_error" in __data:
        _case.FailOnError = __data.get("fail_on_error")
    elif "fail_on_error" in os.environ:
        _case.FailOnError = os.getenv(os.getenv("fail_on_error"))
    else:
        _case.FailOnError = True
    if "exit_on_fail" in __data:
        _case.ExitOnFail = __data.get("exit_on_fail")
    elif "exit_on_fail" in os.environ:
        _case.ExitOnFail = os.getenv(os.getenv("exit_on_fail"))
    else:
        _case.ExitOnFail = True
    if "close_sap_on_cleanup" in __data:
        _case.CloseSAPOnCleanup = __data.get("close_sap_on_cleanup")
    elif "close_sap_on_cleanup" in os.environ:
        _case.CloseSAPOnCleanup = os.getenv(os.getenv("close_sap_on_cleanup"))
    else:
        _case.CloseSAPOnCleanup = True
    if "system" in __data:
        _case.System = __data.get("system")
    elif "system" in os.environ:
        _case.System = os.getenv(os.getenv("system"))
    else:
        _case.System = ""
    _case.Data = __data
    return _case


VKEYS = ["ENTER", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
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


class Strings:
    def transaction_does_not_exist(self) -> tuple:
        return (
            f"Transactie {self.Transaction} bestaat niet", 
            f"Transaction {self.Transaction} does not exist", 
            f"Transaktion {self.Transaction} existiert nicht"
        )


class TextElements(Enum):
    GuiTextField = auto()
    GuiCTextField = auto()
    GuiPasswordField = auto()
    GuiLabel = auto()
    GuiTitlebar = auto()
    GuiStatusbar = auto()
    GuiButton = auto()
    GuiTab = auto()
    GuiShell = auto()
    GuiStatusPane = auto()


class BrowserType(Enum):
    CHROME = auto()
    EDGE = auto()
    FIREFOX = auto()


@dataclass
class Table:
    Id: str 
    Type: str
    TableObject: object
    RowCount: int 
    VisibleRows: int 
    Columns: list[object]
    Rows: list[object]
    Data: list[dict]
