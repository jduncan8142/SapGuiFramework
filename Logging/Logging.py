import logging
from dataclasses import dataclass, field
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


class Logger:
    def __init__(self, config: LoggingConfig) -> None:
        self.log_name = config.LogName
        self.log_path = config.LogPath
        self.log_file = config.LogFilename
        self.verbosity: int = config.LogVerbosity
        self.format: str = config.LogFormat
        self.file_mode: str = config.LogFileMode
        self.stream = config.LogStream
        self.log: logging.Logger = None
        self.formatter: logging.Formatter = None
        self.file_handler: logging.FileHandler = None
        self.stream_handler: logging.StreamHandler = None
        
        if not self.log_file.is_file():
            with open(self.log_file, "w") as f:
                pass
        
        # Create custom logging level for screenshots
        SCREENSHOT_LEVELV_NUM = 25 
        logging.addLevelName(SCREENSHOT_LEVELV_NUM, "SHOT")
        def shot(self, message, *args, **kws):
            if self.isEnabledFor(SCREENSHOT_LEVELV_NUM):
                # Yes, logger takes its '*args' as 'args'.
                self._log(SCREENSHOT_LEVELV_NUM, message, args, **kws)
        logging.Logger.shot = shot

        # Create custom logging level for status
        STATUS_LEVELV_NUM = 55 
        logging.addLevelName(STATUS_LEVELV_NUM, "STATUS")
        def status(self, message, *args, **kws):
            if self.isEnabledFor(STATUS_LEVELV_NUM):
                # Yes, logger takes its '*args' as 'args'.
                self._log(STATUS_LEVELV_NUM, message, args, **kws)
        logging.Logger.status = status

        # Create custom logging level for documentation
        DOCUMENTATION_LEVELV_NUM = 60 
        logging.addLevelName(DOCUMENTATION_LEVELV_NUM, "DOCUMENTATION")
        def documentation(self, message, *args, **kws):
            if self.isEnabledFor(DOCUMENTATION_LEVELV_NUM):
                # Yes, logger takes its '*args' as 'args'.
                self._log(DOCUMENTATION_LEVELV_NUM, message, args, **kws)
        logging.Logger.documentation = documentation

        self.log = logging.getLogger(self.log_name)
        self.formatter = logging.Formatter(self.format)
        self.file_handler = logging.FileHandler(self.log_file, mode=self.file_mode)
        self.file_handler.setFormatter(self.formatter)
        if self.stream:
            self.stream_handler = logging.StreamHandler()
            self.stream_handler.setFormatter(self.formatter)
        match self.verbosity:
            case 5:
                self.log.setLevel(10)
                self.file_handler.setLevel(10)
                if self.stream:
                    self.stream_handler.setLevel(10)
            case 4:
                self.log.setLevel(20)
                self.file_handler.setLevel(20)
                if self.stream:
                    self.stream_handler.setLevel(20)
            case 3:
                self.log.setLevel(25)
                self.file_handler.setLevel(25)
                if self.stream:
                    self.stream_handler.setLevel(30)
            case 2:
                self.log.setLevel(25)
                self.file_handler.setLevel(25)
                if self.stream:
                    self.stream_handler.setLevel(40)
            case 1:
                self.log.setLevel(25)
                self.file_handler.setLevel(25)
                if self.stream:
                    self.stream_handler.setLevel(50)
            case _:
                self.log.setLevel(25)
                self.file_handler.setLevel(25)
                if self.stream:
                    self.stream_handler.setLevel(90)
        self.log.addHandler(self.file_handler)
        if self.stream:
            self.log.addHandler(self.stream_handler)
