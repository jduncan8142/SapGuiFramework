import win32com.client
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class BaseElement:
    Instance: Optional[win32com.client.CDispatch] = None
    Id: Optional[str] = None
    Name: Optional[str] = None
    Text: Optional[str] = None
    ScreenLeft: Optional[int] = None
    ScreenTop: Optional[int] = None
    Handle: Optional[str] = None
    Left: Optional[int] = None
    Top: Optional[int] = None
    Height: Optional[int] = None
    Width: Optional[int] = None
    Tooltip: Optional[str] = None
    DefaultTooltip: Optional[str] = None
    IconName: Optional[str] = None
    Key: Optional[str] = None
    Changeable: Optional[bool] = None
    ContainerType: Optional[bool] = None


@dataclass 
class GuiStatusPane(BaseElement):
    def default_type() -> dict:
        return {"id": 43, "value": "GuiStatusPane"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiStatusbar(BaseElement):
    def default_type() -> dict:
        return {"id": 103, "value": "GuiStatusbar"}
    
    Type: dict = field(default_factory=default_type)
    MessageId: Optional[str] = None
    MessageNumber: Optional[str] = None
    MessageType: Optional[str] = None
    Pane0: Optional[GuiStatusPane] = None
    Pane1: Optional[GuiStatusPane] = None
    Pane2: Optional[GuiStatusPane] = None
    Pane3: Optional[GuiStatusPane] = None
    Pane4: Optional[GuiStatusPane] = None
    Pane5: Optional[GuiStatusPane] = None
    Pane6: Optional[GuiStatusPane] = None

@dataclass
class GuiMenubar(BaseElement):
    def default_type() -> dict:
        return {"id": 111, "value": "GuiMenubar"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiMenu(BaseElement):
    def default_type() -> dict:
        return {"id": 110, "value": "GuiMenu"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiToolbar(BaseElement):
    def default_type() -> dict:
        return {"id": 101, "value": "GuiToolbar"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiButton(BaseElement):
    def default_type() -> dict:
        return {"id": 40, "value": "GuiButton"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiOkCodeField(BaseElement):
    def default_type() -> dict:
        return {"id": 35, "value": "GuiOkCodeField"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiTitlebar(BaseElement):
    def default_type() -> dict:
        return {"id": 102, "value": "GuiTitlebar"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiUserArea(BaseElement):
    def default_type() -> dict:
        return {"id": 74, "value": "GuiUserArea"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiSimpleContainer(BaseElement):
    def default_type() -> dict:
        return {"id": 71, "value": "GuiSimpleContainer"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiTabStrip(BaseElement):
    def default_type() -> dict:
        return {"id": 90, "value": "GuiTabStrip"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiTab(BaseElement):
    def default_type() -> dict:
        return {"id": 91, "value": "GuiTab"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiScrollContainer(BaseElement):
    def default_type() -> dict:
        return {"id": 72, "value": "GuiScrollContainer"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiTextField(BaseElement):
    def default_type() -> dict:
        return {"id": 31, "value": "GuiTextField"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiCTextField(BaseElement):
    def default_type() -> dict:
        return {"id": 32, "value": "GuiCTextField"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiCheckBox(BaseElement):
    def default_type() -> dict:
        return {"id": 42, "value": "GuiCheckBox"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiComboBox(BaseElement):
    def default_type() -> dict:
        return {"id": 34, "value": "GuiComboBox"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiLabel(BaseElement):
    def default_type() -> dict:
        return {"id": 30, "value": "GuiLabel"}
    
    Type: dict = field(default_factory=default_type)


@dataclass
class GuiMainWindow(BaseElement):
    def default_type() -> dict:
        return {"id": 21, "value": "GuiMainWindow"}
    
    Type: dict = field(default_factory=default_type)
