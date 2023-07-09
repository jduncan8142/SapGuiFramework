from dataclasses import dataclass, field
from typing import Optional
from Flow.Results import ResultStep


@dataclass
class Step:
    def default_status() -> ResultStep:
        return ResultStep()
    
    Action: str
    ElementId: str = field(default_factory=str)
    Args: list = field(default_factory=list)
    Kwargs: dict = field(default_factory=dict)
    
    Name: str = field(default_factory=str)
    Description: str = field(default_factory=str)
    
    ApplicationServer: Optional[str] = None
    Language: Optional[str]= None
    Program: Optional[str] = None
    ResponseTime: Optional[float] = None
    RoundTrips: Optional[int] = None
    ScreenNumber: Optional[str] = None
    SystemName: Optional[str] = None
    SystemNumber: Optional[int] = None
    SystemSessionId: Optional[str] = None
    Transaction: Optional[str] = None
    User: Optional[str] = None
    
    PyCode: Optional[str] = field(default_factory=str)
    
    Status: ResultStep = field(default_factory=default_status)
    
    def __repr__(self) -> str:
        return f"""class Step<\
        Name: {self.Name}, \
        Description: {self.Description}, \
        Action: {self.Action}, \
        ElementId: {self.ElementId}, \
        Args: {self.Args}, \
        Kwargs: {self.Kwargs}, \
        ApplicationServer: {self.ApplicationServer}, \
        Language: {self.Language}, \
        Program: {self.Program}, \
        ResponseTime: {self.ResponseTime}, \
        RoundTrips: {self.RoundTrips}, \
        ScreenNumber: {self.ScreenNumber}, \
        SystemName: {self.SystemName}, \
        SystemNumber: {self.SystemNumber}, \
        SystemSessionId: {self.SystemSessionId}, \
        Transaction: {self.Transaction}, \
        User: {self.User}, \
        PyCode: {self.PyCode}, \
        Status: {self.Status}\
        >"""
