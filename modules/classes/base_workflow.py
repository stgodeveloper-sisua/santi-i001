# file_name:   base_workflow.py
# created_on:  2023-04-25 ; sisua.developer
# modified_on: 2023-04-25 ; sisua.developer


import os
import sys
sys.path.append(os.path.abspath(
                    os.path.dirname(
                    os.path.dirname(
                    os.path.abspath(__file__)))))
                    

class WorkflowBase():
    def __init__(self, config:dict):   
        self.config              = config
        self.environment         = config["METADATA"]["ENVIRONMENT"].upper()
        self.config_env          = config[self.environment]
        self.config_global       = config["GLOBAL"]
        self.config_fmw          = config["FRAMEWORK"]        
        self.process_data        = self.config_fmw["PROCESS_DATA"]
        self.config_emails       = self.config["EMAIL"]
        self.be_info             = dict()
        self.process_code        = config["METADATA"]["PROCESS_CODE"].upper()
        self.process_name        = self.config["METADATA"]["PROCESS_NAME"]
        if "EXCEL_VISIBLE" in self.config_env:
            self.excel_visible = self.config_env["EXCEL_VISIBLE"]
        else:
            self.excel_visible = False
