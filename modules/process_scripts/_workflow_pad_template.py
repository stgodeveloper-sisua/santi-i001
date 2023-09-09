# file_name:   _workflow_pad_template.py
# created_on:  2023-05-08 ; sisua.developer
# modified_on: 2023-05-08 ; sisua.developer

import os
import sys
import logging
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # move to modules 
from _fmw.fmw_utils import *
from _fmw.fmw_classes import BusinessException
from classes.base_power_automate import PowerAutomate

# MANDATORY PARAMETERS TO ADD INTO CONFIG FILE
"""
  "GLOBAL": {
    "PAD_CONFIG_FILE": "{PROCESS_DATA}\\pad\\pad_config.xlsx",  // PAD flow: remove it if process doesn't use PAD
    "TRIGGER_EMAIL_SUBJECT": "EJECUTAR PROCESO [{0}]", // 0: PROCESS_CODE  // PAD flow
    "FLOW_FINISHED_LOG_FILE_PATH": "{PROCESS_DATA}\\pad\\pad_flow_finished.log",  // PAD flow: remove it if process doesn't use PAD
    "PAD_FLOW_EXECUTION_LIMIT": 600 // PAD flow: in minutes (default 600 mins (10h))
  },
  "DEV": {
    "CREDENTIALS_FILE_PATH": "C:\\Users\\diazv3\\MetLife\\Acelerar Transformación - OPS - Documentos\\Metlife\\credenciales\\credenciales_bots.xlsx",
    "GET_CREDS_SCRIPT_PATH": "C:\\Users\\diazv3\\MetLife\\Acelerar Transformación - OPS - Documentos\\librerias\\credentials_utils\\get_credential_powershell.py",
    "PAD_TRIGGER_RECIPIENTS": ["sisua.developer@metlifeexternos.cl"],
    "CREDENTIALS": ""
  },
"""

class WorkflowPADTemplate(PowerAutomate):
    def __init__(self, config:dict):
        PowerAutomate.__init__(self, config=config) 
        self.state_name         = "WorkflowPADTemplate"  # Change with the class name 
        self.trigger_subject    = self.config_global["TRIGGER_EMAIL_SUBJECT"].upper().format(self.process_code)

    def add_process_config_params(self):
        # Add process parameters to give to the PAD flow
        self.pad_config["CREDENTIALS"]      = "default"
        return 0
    
    def run_workflow(self):
        logging.info(f"----- Starting state: {self.state_name} -----")
        try: # Add workflow in try block bellow
            self.build_pad_config()        
            self.send_email_trigger_pad()
            self.wait_until_flow_finishes()
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
    state = WorkflowPADTemplate(config=config)
    state.build_pad_config()  # for test build pad config only
    # state.run_workflow()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     