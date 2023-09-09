# file_name:   send_exe_report.py
# created_on:  2023-03-28 ; sisua.developer
# modified_on: 2023-03-29 ; sisua.developer

import os
import sys
import logging
import datetime as dt
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))             
from _fmw.fmw_utils import *
from _fmw.send_email_utils import read_recipients_file, send_email
from _fmw.fmw_classes import BusinessException
from process_scripts._base_process_class import ProcessBase


class SendExeReport(ProcessBase):
    def __init__(self, config:dict):
        ProcessBase.__init__(self, config=config)
        self.state_name    = "SendExeReport"  # Change with the class name
        # These variables already exist by inheritance
        # self.config              = config
        # self.environment         = config["METADATA"]["ENVIRONMENT"].upper()
        # self.config_env          = config[self.environment]
        # self.config_global       = config["GLOBAL"]
        # self.config_fmw          = config["FRAMEWORK"]        
        # self.process_data        = self.config_fmw["PROCESS_DATA"]
        # self.config_emails       = self.config["EMAIL"]
        # self.be_info             = dict()
        # self.process_code        = config["METADATA"]["PROCESS_CODE"].upper()
        self.email_type    = "EXECUTION_REPORT"  
        self.state_success = True
        self.robot_stop    = False
        # Set up now date
        self.now_string = dt.datetime.now().strftime("%d-%m-%y")
        # Get config parameters
        self.recipients_file = self.config_emails["RECIPIENTS_PATH"]
        self.email_wrapper   = self.config_emails["WRAPPER_FILE"]

    def send_execution_report(self): # Alter to fit your process
        logging.info("----- Sending execution report -----")
        # Get body, subject, type
        subject    = self.config_emails[self.email_type]["SUBJECT"].format(self.process_code, self.now_string)
        body_file  = self.config_emails[self.email_type]["BODY_FILE"]
        # Set up outlook parameters
        attachments = []
        body_fields = [self.process_name]
        to,cc,bcc   = read_recipients_file(recipients_file_path=self.recipients_file,
                                           environment=self.environment,
                                           mail_type=self.email_type,
                                           sheet="base")
        # Send email
        send_email(to=to,cc=cc,bcc=bcc,
                   environment=self.environment,
                   subject=subject,
                   body_file=body_file,
                   body_fields=body_fields,
                   wrapper_file=self.email_wrapper,
                   attachments=attachments)
        
    def run_workflow(self):
        logging.info(f"----- Starting state: {self.state_name} -----")
        try: # Add workflow in try block bellow
            self.send_execution_report()
        except Exception as error:
            raise error  
        logging.info(f"Finished state: {self.state_name}") 
        return True
    

if __name__ == "__main__":
    start_logging(logs_level="DEBUG", show_console_logs=True, save_logs=False)
    config = read_config()
    logging.info(f"--------- Script {os.path.basename(__file__)} started ---------")    
    state = SendExeReport(config=config)
    state.run_workflow()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     