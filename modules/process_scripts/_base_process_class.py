# file_name:   _base_process_class.py
# created_on:  2023-08-17 ; sisua.developer
# modified_on: 2023-08-17 ; sisua.developer

import os
import sys
import logging
import pandas as pd
import datetime as dt

sys.path.append(os.path.abspath(
                    os.path.dirname(
                    os.path.dirname(
                    os.path.abspath(__file__)))))
from classes.base_workflow import WorkflowBase
from _fmw.fmw_classes import BusinessException
from classes.robot_date import RobotDate
from _fmw.fmw_utils import *


class ProcessBase(WorkflowBase):
    def __init__(self, config:dict):        
        WorkflowBase.__init__(self, config=config)
        self.state_name      = type(self).__name__       # get the class name, do not change   
        self.execution_date  = RobotDate(config=config)  # to shift execution date in dev or re-executions, do not change
        logging.debug(f"Execution date: {self.execution_date}")
        # processing dates: change depending on the logic of the process (add more if needed)
        self.day_offset      = self.config_global["PROCESSING_DAY_OFFSET"]
        self.process_input_file = self.config_global["PROCESS_INPUT_FILE"]
        self.processing_date = RobotDate(config=config, delta_day=self.day_offset, delta_month=0, delta_year=0) 
        # self.excel_utils = Win32Excel(visible=self.excel_visible, kill_excel_before=True)s
        # Get process parameters from config
        #self.template_parameter_1 = self.config_env["ENV_PARAM_1"]
        #self.template_parameter_2 = self.config_global["GLOBAL_PARAM_1"]
    
    def _build_business_exception(self, error:str): # template for building business exceptions
        # Set up now date
        now_string = dt.datetime.now().strftime("%d-%m-%y")
        # Get config parameters
        be_mail_type   = "BE_REPORT" # Replace BE_REPORT with the BE_NAME you add to config
        be_config_dict = self.config_emails[be_mail_type]
        be_body_file   = be_config_dict["BODY_FILE"]
        process_name   = self.config["METADATA"]["PROCESS_NAME"]
        process_code   = self.config["METADATA"]["PROCESS_CODE"]
        be_subject     = be_config_dict["SUBJECT"].format(process_code, now_string)
        # Set up email parameters
        attachments = [] # add logic if one or more BE needs attachments
        body_fields = [process_name, error]
        # Set up business exception object, do not change
        self.be_info = {
            "BODY_FILE": be_body_file,
            "SUBJECT": be_subject,
            "MAIL_TYPE": be_mail_type,
            "ATTACHMENTS": attachments,
            "BODY_FIELDS": body_fields
        }
        
    def close_excel(self
                    , workbook
                    , save:bool=True):
        if save:
            logging.info(f"Saving workbook {workbook.Name}")
            workbook.Save()
        logging.info(f"Closing workbook {workbook.Name}")
        workbook.Close()

    def process_base_method_1(self):
        logging.info("This is a test method of the process_base class")
        pass

    def run_workflow(self):
        logging.info(f"----- Starting state: {self.state_name} -----")
        try: # Add workflow in try block bellow
            self.process_base_method_1()
        except BusinessException as error:
            self._build_business_exception(error)
            raise error
        except Exception as error:
            raise error  
        logging.info(f"Finished state: {self.state_name}")


if __name__ == "__main__":
    start_logging(logs_level="DEBUG", show_console_logs=True, save_logs=False)
    config = read_config()
    logging.info(f"--------- Script {os.path.basename(__file__)} started ---------")    
    process_base = ProcessBase(config=config)
    process_base.process_base_method_1()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     