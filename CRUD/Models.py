from dataclasses import dataclass, field
from typing import Any, Optional
from pathlib import Path


@dataclass
class LoggingConfig:
    def default_format_factory() -> str:
        return "%(asctime)s|%(levelname)s|%(message)s"
    
    def default_filemode_factory() -> str:
        return "a"
    
    def default_stream_factory() -> bool:
        return False
    
    def default_verbosity_factory() -> int:
        return 4
    
    def default_filepath_factory() -> Path:
        return Path("C:\\temp")
    
    def default_name_factory() -> str:
        return "SapGuiFramework"
    
    def default_filename_factory() -> Path:
        return Path(
            LoggingConfig.default_filepath_factory(), 
            "SapGuiFramework.log"
        )
        
    LogName: str = field(default_factory=default_name_factory)
    LogFilename: Path = field(default_factory=default_filename_factory)
    LogPath: Path = field(default_factory=default_filepath_factory)
    LogVerbosity: int = field(default_factory=default_verbosity_factory)
    LogFormat: str = field(default_factory=default_format_factory)
    LogFileMode: str = field(default_factory=default_filemode_factory)
    LogStream: bool = field(default_factory=default_stream_factory)
