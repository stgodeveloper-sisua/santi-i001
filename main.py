# file_name:   main.py
# created_on:  2023-03-29 ; vicente.diaz
# modified_on: 2023-07-13 ; vicente.diaz ; fixed build system_exception report


import time as tm
import logging
import pythoncom
import traceback
import datetime as dt
from threading import Thread
from modules._fmw.fmw_utils import *
from modules._fmw.send_exceptions_emails import ExceptionEmails
# Import process scripts below
from modules.process_scripts._workflow_template import WorkflowTemplate
from modules.process_scripts.send_exe_report import SendExeReport


class MainProcess(Thread):  
    def __init__(self, state_idx:int=1):  
        super().__init__()
        self.daemon                 = True
        self._config                = read_config()
        self.status                 = "STOPPED"
        self.state_idx              = state_idx
        self._environment           = self._config["METADATA"]["ENVIRONMENT"].upper()
        self._config_env            = self._config[self._environment] 
        self._cfg_max_tries         = self._config["FRAMEWORK"]["MAX_TRIES"]
        self._kill_processes        = self._config["FRAMEWORK"]["KILL_PROCESSES"]
        self._kill_proces_list      = self._config["FRAMEWORK"]["LIST_PROCESESS_TO_KILL"]
        self._tries_counter         = 1
        self._running               = False
        self.current_workflow       = None
        self._exception             = None

    def run(self):
            logging.info("--------- Starting MainProcess execution ---------")
            pythoncom.CoInitialize()  # check description
            self._running = True
            t1 = tm.time()
            while self._running:
                self.status = "_running"
                try:
                    # Start invocation of state scripts
                    if self.state_idx == 1:
                        self.current_workflow = WorkflowTemplate(self._config)         
                    elif self.state_idx == 2:
                        self.current_workflow = WorkflowTemplate(self._config)
                    # elif self.state_idx == 3:
                    #     self.current_workflow = SendExeReport(self._config)
                    else:
                        self._running = False
                        self.status   = "SUCCESS"
                        if not self.current_workflow:
                            logging.warning(f"State {self.state_idx} not implemented.")
                        logging.info(f"Execution finished")
                        logging.info(f"Elapsed time (main.py):  {dt.timedelta(seconds=tm.time()-t1)}")
                        break
                    # Run state workflow
                    logging.info(f"----- Starting state: {self.state_idx} ------")
                    self.current_workflow.run_workflow()
                    self.state_idx += 1
                    self._tries_counter = 1
                except Exception as error:
                    logging.error(f"Exception found: Type: {type(error)} | Msg: {error}")
                    self.status   = "FAILED"
                    if self._kill_processes:
                        kill_processes(processes=self._kill_proces_list)
                    logging.warning(f"Try: {self._tries_counter}/{self._cfg_max_tries}")
                    if "BusinessException" in str(type(error)).split(".")[-1]:
                        self._running   = False
                        self.status     = "WARNING"
                        self._exception = error
                    elif self._tries_counter >= self._cfg_max_tries:
                        logging.warning("Max tries reached! Throwing exception")
                        self._exception = traceback.format_exc()
                        self._running = False
                    else:
                        logging.warning(f"Retrying state_idx {self.state_idx}")
                        del self.current_workflow
                        self._tries_counter += 1
            # Check if execution ended with Exception
            if self.status == "FAILED" or self.status == "WARNING":
                self.notify_exception()
            # end run function

    def notify_exception(self):
        logging.info("--- Starting notify_exception ---")
        logging.error(self._exception)
        error_emails = ExceptionEmails(self._config)
        # Determine if business exception
        if "BusinessException" in str(type(self._exception)).split(".")[-1]:
            # Send Business Exception
            be_info = self.current_workflow.be_info
            error_emails.send_business_exception(be_info=be_info)
        else: # System exception
            try:
                # Get all lines in traceback
                lines    = self._exception.strip().split("\n")
                # Build traceback dataframe
                exc_tree = list()
                for idx, line in enumerate(lines):
                    if "File " in line.strip()[:min(5, len(lines))]:
                        file      = line.split('"')[1].replace(os.path.dirname(__file__)+"\\", "")
                        line_idx  = int(line.split('line ')[-1].split(",")[0])
                        func      = line.split('in ')[-1] + "()"
                        line_code = lines[idx+1].strip()
                        exc_info  = {"file": file, "line_idx": line_idx, "function": func, "code": line_code}
                        exc_tree .append(exc_info)
                        logging.info(exc_info)
                exception_df = pd.DataFrame(exc_tree)
                exc_info = lines[-1].split(": ")
                exc_source = exc_info[0]
                exc_msg    = " : ".join(exc_info[1:])
            except Exception as error:
                logging.warning("An error occured while trying to set up system exception, sending simplified version")
                exception_df = pd.DataFrame()
                exc_source   = "CHECK LOGS"
                exc_msg      = self._exception.strip()
            logging.info(f"Exception source: {exc_source} | Msg: {exc_msg}")
            # Send system exception
            error_emails.send_system_exception(exc_message=exc_msg, exc_source=exc_source, exception_df=exception_df)
            error_emails.send_user_system_exception()
        return True

    
if __name__ == "__main__":
    start_logging(logs_level="INFO", show_console_logs=True, save_logs=False)
    logging.info("--------- Script main.py started ---------")
    robot_instance = MainProcess(state_idx=1)
    robot_instance.run()
    logging.info("--------- Script main.py finished ---------")