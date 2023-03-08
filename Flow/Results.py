from enum import StrEnum, auto
from dataclasses import dataclass, field

class Result(StrEnum):
    PASS = auto()
    FAIL = auto()
    WARN = auto()


@dataclass
class ResultCase:    
    Result: Result = None
    FailedSteps: list = field(default_factory=list)
    FailedScreenShots: list = field(default_factory=list)
    PassedSteps: list = field(default_factory=list)
    PassedScreenShots: list = field(default_factory=list)


@dataclass
class ResultStep:
    Result: Result = None
    Error: str = None
