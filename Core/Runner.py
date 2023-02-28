from Flow.Actions import Step
import inspect


class Runner:
    def __init__(self) -> None:
        pass    
    
    def exec_step(self, step_id: int, step: Step) -> None:
        self.current_step = step
        step.Action()
        if self.current_step.Transaction != self.current_transaction:
            self.start_transaction(transaction=step.Transaction)
        try:
            self.collect_step_meta_data()
        except Exception as err:
            self.logger.log.warning(f"Unhandled Exception: {err} while collecting step metadata for case: {self.case.Name} <> Step: {step_id}")
        # try:
        #     print(f"Evaluating => {self.current_step.PyCode}")
        #     # r = eval(self.current_step.PyCode)
        #     # print(f"Result > {r}")
        #     # exec(self.current_step.PyCode)
        #     self.step_pass(msg=f"{self.current_step.Description}", ss_name=self.current_step.Name)
        # except Exception as err:
        #     self.step_fail(msg=f"{self.current_step.Description}", ss_name=self.current_step.Name)
        #     self.handle_unknown_exception(f"Unhandled exception: {err} while executing {self.current_step.PyCode}", ss_name="execute_step::UnhandledException")
    
    def exec_case(self) -> None:
        self.collect_case_meta_data()
        with open(self.case.StepsFile, "a") as sf:
            sf.writelines([f"{i.PyCode}\n" for i in self.case.Steps])
        for __step_id, __step in enumerate(self.case.Steps):
            self.execute_step(__step_id, __step) 
