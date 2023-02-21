import string
import yaml
from typing import Optional
from Flow.Actions import Step


class Parser:
    def __init__(self) -> None:
        pass
    
    def parse_if_condition(self, condition) -> str:
        my_condition: str = condition
        return my_condition

    def parse_value(self, value) -> str:
        my_value_final_list: list = []
        my_value_list: list = value.split(".")
        for i, v in enumerate(my_value_list):
            if v == "Case":
                my_value_final_list.append("self.case")
            elif v == "Data":
                my_value_final_list.append(".Data")
            elif v == "System":
                my_value_final_list.append(".System")
            elif v in string.digits:
                my_value_final_list.append(f"{v}")
            elif my_value_list[i - 1] in ("Case", "Data"):
                my_value_final_list.append(f"['{v}']")
            elif "(" in list(v):
                __tmp = v.split("(")
                my_value_final_list.append(f"self.session.{__tmp[0]}('{self.auto_complete_element_id(id=__tmp[1].replace(')', ''))}')")
            else:
                my_value_final_list.append(f"{v}")
        return "".join(my_value_final_list)


    def generate_py_code(self, step: Step, loops: Optional[dict] = None) -> tuple[str, dict]:
        my_code: str = None
        my_loops: dict = loops if loops is not None else {'while': 0, 'for': 0, 'if': 0, 'else': 0, 'elif': 0, 'range': 0, 'try': 0, 'except': 0}
        my_indent: int = 0
        for k, v in my_loops.items():
            if v != 0:
                my_indent = my_indent + v
        match step.Action:
            case "open_connection":
                my_code = f"self.session.open_connection({self.parse_value(step.Args[0])})"
            case "start_transaction":
                my_code = f"self.session.start_transaction('{step.Args[0]}')"
            case "documentation":
                my_code = f"self.session.documentation('{step.Args[0]}')"
            case "input_text":
                my_code = f"self.session.input_text(id='{self.auto_complete_element_id(step.Args[0])}', text={self.parse_value(step.Args[1])})"
            case "set_variable":
                var_and_type = step.Args[0].split(":")
                if len(var_and_type) == 2:
                    my_code = f"{var_and_type[0]}: {var_and_type[1]} = {self.parse_value(step.Args[1])}"
                if len(var_and_type) == 1:
                    my_code = f"{var_and_type[0]} = {self.parse_value(step.Args[1])}"
            case "set_v_scrollbar":
                my_code = f"self.session.set_v_scrollbar(id='{self.auto_complete_element_id(step.Args[0])}', pos={self.parse_value(step.Args[1])})"
            case "start_while":
                my_loops['while'] += 1
                my_code = f"while {step.Args[0]}:"
            case "end_while":
                my_loops['while'] -= 1
                my_code = f""
            case "start_range":
                my_loops['range'] += 1
                my_code = f"for i in range({self.parse_value(step.Args[0])}, {self.parse_value(step.Args[1])}):"
            case "end_range":
                my_loops['range'] -= 1
                my_code = f""
            case "start_try":
                my_loops['try'] += 1
                my_code = f"try:"
            case "end_try":
                my_loops['try'] -= 1
                my_code = f""
            case "start_except":
                my_loops['except'] += 1
                if step.Args[0]:
                    my_code = f"except {step.Args[0]} as err:"
                else:
                    my_code = f"except Exception as err:"
            case "end_except":
                my_loops['except'] -= 1
                my_code = f""
            case "start_if":
                my_loops['if'] += 1
                my_code = f"if {self.parse_if_condition(step.Args[0])}:"
            case "end_if":
                my_loops['if'] -= 1
                my_code = f""
            case "start_else":
                my_loops['else'] += 1
                my_loops['if'] -= 1
                my_code = f"else:"
            case "end_else":
                my_loops['else'] -= 1
                my_loops['if'] += 1
                my_code = f""
            case "start_elif":
                my_loops['elif'] += 1
                my_loops['if'] -= 1
                my_code = f"elif {self.parse_if_condition(step.Args[0])}:"
            case "end_elif":
                my_loops['elif'] -= 1
                my_loops['if'] += 1
                my_code = f""
            case "update_var":
                my_code = f"{step.Args[0]}{step.Args[1]}"
            case "enter": 
                my_code = f"self.session.enter()"
            case "click_element":
                my_code = f"self.session.click_element(id='{step.Args[0]}')"
            case "wait_for_element":
                my_code = f"self.session.wait_for_element(id='{step.Args[0]}')"
            case "save":
                my_code = f"self.session.save()"
            case "take_screenshot":
                try:
                    if step.Args[0] and step.Args[1]:
                        my_code = f"self.session.take_screenshot(filename='{step.Args[0]}', id='{step.Args[1]}')"
                    elif step.Args[0]:
                        my_code = f"self.session.take_screenshot(filename='{step.Args[0]}')"
                    else:
                        my_code = f"self.session.take_screenshot()"
                except Exception as err:
                    pass
                    # print(f"ERROR > {err}")
            case "wait":
                my_code = f"self.session.wait({step.Args[0]})"
        return (my_code, my_loops)
    
    def parse_step_from_string(self, step_string: str, loops: Optional[dict] = None) -> tuple[Step, dict]:
        __loops: dict = loops
        self.current_step = Step()
        ss: list = step_string.split("|")
        self.current_step.Action = ss[0]
        if len(ss) >= 2: self.current_step.Args = ss[1].split(";")
        if len(ss) >= 3: self.current_step.Description = ss[2]
        if len(ss) >= 4: self.current_step.FailOnError = ss[3]
        if len(ss) >= 5: self.current_step.Name = ss[4]
        if len(ss) >= 6: self.current_step.ScreenShotOnFail = ss[5]
        if len(ss) >= 7: self.current_step.ScreenShotOnPass = ss[6]
        self.collect_step_meta_data()
        self.current_step.PyCode, __loops = self.generate_py_code(step=self.current_step, loops=__loops)
        self.current_step.Status = ResultDetails()
        return self.current_step, __loops
    
    def load_case_from_yaml(self, file_path: Path) -> None:
        __loops: dict = None
        cff = yaml.safe_load(file_path.open())
        if 'Name' in cff['Case']: self.case.Name = cff['Case']['Name']  
        if 'Description' in cff['Case']: self.case.Description = cff['Case']['Description']
        if 'BusinessProcessOwner' in cff['Case']: self.case.BusinessProcessOwner = cff['Case']['BusinessProcessOwner']
        if 'ITOwner' in cff['Case']: self.case.ITOwner = cff['Case']['ITOwner']        
        if 'DocumentationLink' in cff['Case']: self.case.DocumentationLink = cff['Case']['DocumentationLink'] 
        if 'BasePath' in cff['Case']: self.case.BasePath = cff['Case']['BasePath']
        if 'LogConfig' in cff['Case']: self.case.LogConfig = cff['Case']['LogConfig']
        if 'DateFormat' in cff['Case']: self.case.DateFormat = cff['Case']['DateFormat']
        if 'ExplicitWait' in cff['Case']: self.case.ExplicitWait = cff['Case']['ExplicitWait']
        if 'ScreenShotOnPass' in cff['Case']: self.case.ScreenShotOnPass = cff['Case']['ScreenShotOnPass']
        if 'ScreenShotOnFail' in cff['Case']: self.case.ScreenShotOnFail = cff['Case']['ScreenShotOnFail']
        if 'FailOnError' in cff['Case']: self.case.FailOnError = cff['Case']['FailOnError']
        if 'ExitOnError' in cff['Case']: self.case.ExitOnError = cff['Case']['ExitOnError']
        if 'CloseSAPOnCleanup' in cff['Case']: self.case.CloseSAPOnCleanup = cff['Case']['CloseSAPOnCleanup']
        if 'System' in cff['Case']: self.case.System = cff['Case']['System']
        
        if 'Steps' in cff['Case']:
            string_steps = [i for i in cff['Case']['Steps'].split("\n")]
            if len(string_steps) > 0:
                for __step in string_steps:
                    __s, __loops = self.parse_step_from_string(step_string=__step, loops=__loops)
                    self.case.Steps.append(__s)
        if 'Data' in cff['Case']: self.case.Data = cff['Case']['Data']
        self.case.Status: Optional[ResultDetails] = ResultDetails()
        self.case.StepsFile: str = r"C:\Users\duncan\github\SapGuiFramework\steps.py"
        self.case.Error: Optional[str] = None
        self.collect_case_meta_data()
