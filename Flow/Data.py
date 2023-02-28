from dataclasses import dataclass, field
from typing import Optional
from pathlib import Path
from Utilities import get_main_dir
from enum import Enum
import xml.etree.ElementTree as ET
from Flow.Actions import Step
from Flow.Results import ResultCase
from CRUD.Models import LoggingConfig
from datetime import datetime

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
    def empty_string() -> str:
        return ""

    def default_case_path() -> Path:
        return get_main_dir()
    
    def default_explicit_wait() -> float:
        return 0.25
    
    def default_date_format() -> str:
        return "%m/%d/%Y"
    
    def empty_list_factory() -> list:
        return []
    
    def default_name() -> str:
        return f"test_{datetime.now().strftime('%m%d%Y_%H%M%S')}"
    
    def default_business_process_owner() -> str:
        return "Business Process Owner"
    
    def default_it_owner() -> str:
        return "Business Process Owner"
    
    def default_log_config() -> LoggingConfig:
       return LoggingConfig()
   
    def default_system() -> str:
        return "1.2 ERP - RQ2"
    
    def default_result() -> ResultCase:
        return ResultCase()
    
    Name: str = field(default_factory=default_name)
    Description: str = field(default_factory=empty_string)
    BusinessProcessOwner: str = field(default_factory=default_business_process_owner)
    ITOwner: str = field(default_factory=default_it_owner)
    DocumentationLink: str = field(default_factory=empty_string)
    CasePath: Path = field(default_factory=default_case_path)
    LogConfig: LoggingConfig = field(default_factory=default_log_config)
    DateFormat: str = field(default_factory=default_date_format)
    ExplicitWait: float = field(default_factory=default_explicit_wait)
    
    ScreenShotOnPass: bool = False
    ScreenShotOnFail: bool = False
    FailOnError: bool = True
    ExitOnFail: bool = True
    CloseSAPOnCleanup: bool = True
    
    System: str = field(default_factory=default_system)
    Steps: list[Step] = field(default_factory=empty_list_factory)
    Data: dict = None
    Status: ResultCase = field(default_factory=default_result)
    
    SapMajorVersion: Optional[int] = None
    SapMinorVersion: Optional[int] = None
    SapPatchLevel: Optional[int] = None
    SapRevision: Optional[int] = None
        

# class VKeys(Enum):
#     ENTER = "ENTER"
#     F1 = "F1"
#     F2 = "F2"
#     F3 = "F3"
#     F4 = "F4"
#     F5 = "F5"
#     F6 = "F6"
#     F7 = "F7"
#     F8 = "F8"
#     F9 = "F9"
#     F10 = "F10"
#     F11 = "F11"
#     F12 = "F12"
#     SHIFT = "SHIFT"
#     CTRL = "CTRL"
#     A = "A"
#     E = "E"
#     F = "F"
#     D = "D"
#     N = "N"
#     O = "O"
#     DEL = "DEL"
#     INS = "INS"
#     ALT = "ALT"
#     BACKSPACE = "BACKSPACE"
#     PAGEUP = "PAGEUP"
#     PAGEDOWN = "PAGEDOWN"
#     G = "G"
#     R = "R"
#     P = "P"
#     B = "B"
#     K = "K"
#     T = "T"
#     Y = "Y"
#     X = "X"
#     C = "C"
#     V = "V"


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
    GuiTextField = "GuiTextField"
    GuiCTextField = "GuiCTextField"
    GuiPasswordField = "GuiPasswordField"
    GuiLabel = "GuiLabel"
    GuiTitlebar = "GuiTitlebar"
    GuiStatusbar = "GuiStatusbar"
    GuiButton = "GuiButton"
    GuiTab = "GuiTab"
    GuiShell = "GuiShell"
    GuiStatusPane = "GuiStatusPane"


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
