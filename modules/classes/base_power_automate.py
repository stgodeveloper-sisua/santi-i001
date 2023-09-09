# file_name:   base_power_automate.py
# created_on:  2023-05-05 ; vicente.diaz
# modified_on: 2023-05-16 ; vicente.diaz ; handle business exceptions

import os
import sys
import time
import logging
import datetime as dt
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # move to modules 
from _fmw.fmw_utils import *
from _fmw.fmw_classes import BusinessException
from _fmw.send_email_utils import send_email
from process_scripts._base_process_class import ProcessBase

# MANDATORY PARAMETERS TO ADD INTO CONFIG FILE
"""
  "GLOBAL": {
    "PAD_CONFIG_FILE": "{PROCESS_DATA}\\pad\\pad_config.xlsx",  // remove it if process doesn't use PAD
    "TRIGGER_EMAIL_SUBJECT": "EJECUTAR PROCESO [{0}]", // 0: PROCESS_CODE
    "FLOW_FINISHED_LOG_FILE_PATH": "{PROCESS_DATA}\\pad\\pad_flow_finished.log",  // remove it if process doesn't use PAD
    "PAD_FLOW_EXECUTION_LIMIT": 600 // in minutes (10 hours max), depends on the PAD flow
  },
  "DEV": {
    "CREDENTIALS_FILE_PATH": "C:\\Users\\diazv3\\MetLife\\Acelerar Transformación - OPS - Documentos\\Metlife\\credenciales\\credenciales_bots.xlsx",
    "GET_CREDS_SCRIPT_PATH": "C:\\Users\\diazv3\\MetLife\\Acelerar Transformación - OPS - Documentos\\librerias\\credentials_utils\\get_credential_powershell.py",
    "PAD_TRIGGER_RECIPIENTS": ["sisua.developer@metlifeexternos.cl"],
    "CREDENTIALS": ""
  },
"""

class PowerAutomate(ProcessBase):
    def __init__(self, config:dict):
        ProcessBase.__init__(self, config=config) 
        self.state_name         = "PowerAutomate"  # Change with the class name 
        self.minutes_limit      = self.config_global["PAD_FLOW_EXECUTION_LIMIT"] 
        self.pad_config_file    = self.config_global["PAD_CONFIG_FILE"]
        self.trigger_subject    = self.config_global["TRIGGER_EMAIL_SUBJECT"].upper().format(self.process_code)
        self.pad_config         = dict()
        # Env attributes
        self.trigger_recipients        = self.config_env["PAD_TRIGGER_RECIPIENTS"]
        self.get_creds_script_original = self.config_env["GET_CREDS_SCRIPT_PATH"]
        self.credentials_file_path     = self.config_env["CREDENTIALS_FILE_PATH"]
        # Internal attributes
        internal_pad_folder        = os.path.join(self.config_fmw["PROCESS_DATA"], "pad")
        self.email_default_file    = os.path.join(internal_pad_folder, "email_pad_body_temp_file.txt")
        self.get_creds_script_temp = os.path.join(internal_pad_folder, "pad_get_credential.py")
        self.flag_creds_file_path  = "%CREDS_FILE_PATH%"  # same as the variable in the GET_CREDS_SCRIPT_PATH
        delete_folder(folder_path=internal_pad_folder)
        create_folder(folder_path=internal_pad_folder)
        # Add default config parameters
        self.pad_config["FLOW_FINISHED_LOG_FILE_PATH"] = self.config_global["FLOW_FINISHED_LOG_FILE_PATH"]
        self.pad_config["GET_CREDS_SCRIPT_PATH"]       = self.get_creds_script_temp

    def create_temp_get_creds(self):
        shutil.copy(src=self.get_creds_script_original, dst=self.get_creds_script_temp)
        with open(self.get_creds_script_temp, "r", encoding="UTF-8") as file:
            new_file = file.read().replace(self.flag_creds_file_path, self.credentials_file_path)
            new_file = new_file.replace("\\", "/")
        with open(self.get_creds_script_temp, "w", encoding="UTF-8") as file:
            file.write(new_file)

    def add_default_parameters(self):
        self.pad_config["FLOW_FINISHED_LOG_FILE_PATH"] = self.config_global["FLOW_FINISHED_LOG_FILE_PATH"]
        self.pad_config["GET_CREDS_SCRIPT_PATH"]       = self.get_creds_script_temp
        return 0
    
    def add_process_config_params(self):
        # Add process parameters to give to the PAD flow
        return 0

    def build_pad_config(self):
        logging.info(f"--- BUILDING PAD CONFIG ---")
        # Run script after this
        self.create_temp_get_creds()
        self.add_default_parameters()
        # Add process parameters to give to the PAD flow
        self.add_process_config_params()
        # Save PAD config into excel file
        logging.info(f"PAD config parameters: {len(self.pad_config)}")
        df_pad_config = pd.DataFrame(data=self.pad_config.items(), columns=["key", "value"])
        save_excel_file(df=df_pad_config, file_path=self.pad_config_file)
        return 0
    
    def send_email_trigger_pad(self):
        logging.info(f"--- SENDING EMAIL TRIGGER PAD ---")
        # Create empty wrapper and email body file
        with open(self.email_default_file, "w") as file:
            file.write("{0}")
        # Send email to trigger the PAD flow
        send_email(to=self.trigger_recipients
                ,cc=[]
                ,bcc=[]
                ,subject=self.trigger_subject
                ,body_file=self.email_default_file
                ,body_fields=[f"{self.pad_config_file}end_config"]
                ,wrapper_file=self.email_default_file
                ,attachments=[])
        return 0
    
    def wait_until_flow_finishes(self):
        seconds_to_wait = int(self.minutes_limit * 60) + 1
        logging.info(f"Flag file: {self.pad_config['FLOW_FINISHED_LOG_FILE_PATH']}")
        logging.info(f"Waiting for the PAD flow to finish ({self.minutes_limit} minutes max) ...")
        t1 = time.time()
        for second in range(seconds_to_wait):
            if (second) % 60 == 0:
                logging.info(f"Minutes waiting: {second//60}")
            pad_flow_finished = os.path.exists(self.pad_config["FLOW_FINISHED_LOG_FILE_PATH"])
            if pad_flow_finished:
                logging.info(f"Flag file found!")
                logging.info(f"Elapsed time (PAD flow): {dt.timedelta(seconds=time.time()-t1)}")
                self.log_pad_flow_logs()
                break
            time.sleep(1) # delay 1 second
        if not pad_flow_finished:
            logging.critical(f"PAD flow not finished!")

    def log_pad_flow_logs(self):
        logging.info(f"Reading PAD flow log file")
        log_file = self.pad_config["FLOW_FINISHED_LOG_FILE_PATH"]
        flow_errors = list()
        flow_be     = list()
        with open(log_file, "r", encoding="utf-16") as file:
            lines = [line for line in file.read().split("\n") if len(line) > 1]
            for line in lines:
                if "[BE]" in line:
                    logging.warning(f"PAD: {line}")
                    flow_be.append(line)
                elif "[ERROR]" in line:
                    logging.error(f"PAD: {line}")
                    flow_errors.append(line)
                else:
                    logging.info(f"PAD: {line}")
        logging.info(f"Read PAD flow log file finished")
        if len(flow_be) > 0:
            error_msg = ";".join(flow_be)
            raise BusinessException(f"{error_msg}")
        if len(flow_errors) > 0:
            error_msg = ";".join(flow_errors)
            raise Exception(f"Errors found in PAD flow: {len(flow_errors)}\n{error_msg}")

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
    state = PowerAutomate(config=config)
    state.build_pad_config()
    # state.run_workflow()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     