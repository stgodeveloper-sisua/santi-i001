# file_name:   send_exe_report.py
# created_on:  2023-03-28 ; sisua.developer
# modified_on: 2023-03-29 ; sisua.developer

import os
import sys
import logging
import datetime as dt
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))             
from _fmw.fmw_utils import *
from classes.base_send_email import SendEmailBase
from process_scripts._base_process_class import ProcessBase


class SendExeReport(ProcessBase, SendEmailBase):
    def __init__(self, config:dict):
        ProcessBase.__init__(self, config=config) 
        SendEmailBase.__init__(self, config=config, mail_type="EXECUTION_REPORT")
        self.state_name = "SendExeReport"      # Change with the class name           

    def run_workflow(self):
        logging.info(f"----- Starting state: {self.state_name} -----")
        self.list_source_files = [] # add attchments
        self.send_email_base()
        
    
if __name__ == "__main__":
    start_logging(logs_level="DEBUG", show_console_logs=True, save_logs=False)
    config = read_config()
    logging.info(f"--------- Script {os.path.basename(__file__)} started ---------")    
    state = SendExeReport(config=config)
    state.run_workflow()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     