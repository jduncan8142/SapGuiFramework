from dataclasses import dataclass, field
from typing import Optional
from Flow.Results import ResultStep


@dataclass
class Step:
    def default_status() -> ResultStep:
        return ResultStep()
    
    def default_name() -> str:
        return ""
    
    def default_description() -> str:
        return ""
    
    def default_args() -> list:
        return []
    
    Action: str = None
    ElementId: str = None
    Args: list = field(default_factory=default_args)
    
    Name: str = field(default_factory=default_name)
    Description: str = field(default_factory=default_description)
    
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
    
    PyCode: Optional[str] = None
    
    Status: ResultStep = field(default_factory=default_status)
