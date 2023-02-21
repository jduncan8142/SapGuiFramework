from enum import Enum
from dataclasses import dataclass, field

class Result(Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    WARN = "WARN"


@dataclass
class ResultCase:
    def empty_list_factory() -> list:
        return []
    
    Result: Result = None
    FailedSteps: list = field(default_factory=empty_list_factory)
    FailedScreenShots: list = field(default_factory=empty_list_factory)
    PassedSteps: list = field(default_factory=empty_list_factory)
    PassedScreenShots: list = field(default_factory=empty_list_factory)


@dataclass
class ResultStep:
    Result: Result = None
    Error: str = None
