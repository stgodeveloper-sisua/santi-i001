# file_name:   _workflow_template.py
# created_on:  2023-08-17 ; sisua.developer
# modified_on: 2023-08-17 ; vicente.diaz ; updated self.state_name attribute

import os
import sys
import logging
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # move to modules 
from _fmw.fmw_utils import *
from _fmw.fmw_classes import BusinessException
from classes.robot_date import RobotDate
from process_scripts._base_process_class import ProcessBase


class WorkflowTemplate(ProcessBase):
    def __init__(self, config:dict):
        ProcessBase.__init__(self, config=config) 
        self.state_name = self.state_name = type(self).__name__ # Get the class name, do not change     
        # workflow parameters 
        self.template_parameter_1 = self.config_env["ENV_PARAM_1"]
        self.template_parameter_2 = self.config_global["GLOBAL_PARAM_1"]
        # These variables already exist by inheritance
        # self.config              = config
        # self.environment         = config["METADATA"]["ENVIRONMENT"].upper()
        # self.config_env          = config[self.environment]
        # self.config_global       = config["GLOBAL"]
        # self.config_fmw          = config["FRAMEWORK"]        
        # self.process_data        = self.config_fmw["PROCESS_DATA"]
        # self.config_emails       = self.config["EMAIL"]
        # self.now                 = dt.datetime.now()

    def script_function_1(self):
        logging.info(f"--- DESCRIPTIVE START FOR SPRIPT 1 ---")
        # Run script after this
        logging.info(f"Reached example step 1 {self.template_parameter_1}")
        logging.info(f"DESCRIPTIVE END FOR SPRIPT 1")
        # raise BusinessException("This is a test business exception")
        return 0
    
    def script_function_2(self):
        logging.info(f"--- DESCRIPTIVE START FOR SPRIPT 2 ---")
        # Run script after this
        # raise BusinessException("Testing business exceptions") # Testing business exception
        logging.info(f"Reached example step 2 {self.template_parameter_2}")
        logging.info(f"DESCRIPTIVE END FOR SPRIPT 2")
        return 0

    def run_workflow(self):
        logging.info(f"----- Starting state: {self.state_name} -----")
        try: # Add workflow in try block bellow
            self.script_function_1()        
            self.script_function_2()
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
    state = WorkflowTemplate(config=config)
    state.run_workflow()
    # state.script_function_1()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     