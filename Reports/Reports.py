from Utilities.Utilities import *


class Reporter:
    def __init__(self, test_dir: str, test_case: Optional[str] = None, output_dir: Optional[str] = "output", data_file: Optional[str] = "data.py", log_file: Optional[str] = None) -> None:
        self.test_dir: str = test_dir
        self.current_dir: str = os.getcwd()
        os.chdir(self.test_dir)
        self.output_dir: str = output_dir
        self.data_file: str = data_file
        # try:
        #     if os.path.isfile(data.py):
        #         pass
        #         # import data
        #     else:
        #         raise FileExistsError(f"Data file: {self.data_file} does not exist")
        # except Exception as e:
        #     raise FileNotFoundError(f"Data file: {self.data_file} was unable to be opened > {e}")
        # self.test_case: str = test_case if test_case is not None else data.test_case
        # self.log_file: str = log_file if log_file is not None else f"{self.test_case}.log"
        # self.log_data = open(os.path.join(self.output_dir, self.log_file)).readlines()
        # for row in self.log_data:
        #     print(row)
