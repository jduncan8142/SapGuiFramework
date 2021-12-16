import logging
import os
from typing import Optional
from Utilities.Utilities import get_main_dir


class Logger:
    def __init__(
            self, 
            log_name: Optional[str] = None, 
            log_path: Optional[str] = None, 
            log_file: Optional[str] = None, 
            verbosity: Optional[int] = None, 
            format: Optional[str] = None, 
            file_mode: Optional[str] = "a") -> None:
        self.log_name: str = log_name
        if self.log_name is None:
            self.log_name = "MyLogs"

        self.log_path: str = log_path
        if log_path is None:
            self.log_path = os.path.join(get_main_dir(), "logs")
        if not os.path.isdir(self.log_path):
            os.mkdir(self.log_path)

        self.log_file: str = log_file 
        if self.log_file is None:
            os.path.join(self.log_path, f"{self.log_name}.log")
        if not os.path.isfile(self.log_file):
            with open(self.log_file, "w") as f:
                pass

        self.verbosity: int = verbosity
        if self.verbosity is None:
            self.verbosity = 4
        self.format: str = format
        if self.format is None:
            self.format = "%(asctime)s|%(levelname)s|%(message)s"
        self.file_mode = file_mode if file_mode is not None else "a"
        
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
        DOUMENTATION_LEVELV_NUM = 60 
        logging.addLevelName(DOUMENTATION_LEVELV_NUM, "DOCUMENTATION")
        def documentation(self, message, *args, **kws):
            if self.isEnabledFor(DOUMENTATION_LEVELV_NUM):
                # Yes, logger takes its '*args' as 'args'.
                self._log(DOUMENTATION_LEVELV_NUM, message, args, **kws)
        logging.Logger.documentation = documentation

        self.log: logging.Logger = logging.getLogger(self.log_file)
        self.formatter: logging.Formatter = logging.Formatter(self.format)
        self.file_handler: logging.FileHandler = logging.FileHandler(self.log_file, mode=self.file_mode)
        self.file_handler.setFormatter(self.formatter)
        self.stream_handler: logging.StreamHandler = logging.StreamHandler()
        self.stream_handler.setFormatter(self.formatter)
        match self.verbosity:
            case 5:
                self.__log.setLevel(10)
                self.file_handler.setLevel(10)
                self.stream_handler.setLevel(10)
            case 4:
                self.__log.setLevel(20)
                self.file_handler.setLevel(20)
                self.stream_handler.setLevel(20)
            case 3:
                self.__log.setLevel(25)
                self.file_handler.setLevel(25)
                self.stream_handler.setLevel(30)
            case 2:
                self.__log.setLevel(25)
                self.file_handler.setLevel(25)
                self.stream_handler.setLevel(40)
            case 1:
                self.__log.setLevel(25)
                self.file_handler.setLevel(25)
                self.stream_handler.setLevel(50)
            case _:
                self.__log.setLevel(25)
                self.file_handler.setLevel(25)
                self.stream_handler.setLevel(90)
        self.__log.addHandler(self.file_handler)
        self.__log.addHandler(self.stream_handler)
