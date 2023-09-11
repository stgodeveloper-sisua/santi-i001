# file_name:   fmw_classes.py
# created_on:  2023-03-17 ; vicente.diaz ; juanpablo.mena
# modified_on: 2023-03-29 ; vicente.diaz ; guillermo.konig

import logging
import datetime as dt


class BusinessException(Exception):
    # Class to raise Business exceptions in the scripts 
    # Example: raise BusinessException("This is a business exception")
    pass


class Worktray():
    # Pending
    pass


class Workflow:
    def __init__(self, config:dict):
        self.state_name    = "Workflow"  # Change with the class name         
        self.config        = config
        self.environment   = self.config["METADATA"]["ENVIRONMENT"].upper()
        self.config_global = self.config["GLOBAL"]
        self.config_emails = self.config["EMAIL"]
        self.config_env    = self.config[self.environment]
        self.be_info       = dict()        
        # Get class parameters from config
        self.template_parameter_1 = self.config_env["ENV_PARAM_1"]
        self.template_parameter_2 = self.config_global["GLOBAL_PARAM_1"]  

    def _build_business_exception(self, error:str): # template for building business exceptions
        # Set up now date
        now_string = dt.datetime.now().strftime("%d-%m-%y")
        # Get config parameters
        be_config_dict = self.config_emails["BE_REPORT"] # Replace BE_REPORT with the BE_NAME you add to config
        be_body_file   = be_config_dict["BODY_FILE"]
        be_mail_type   = be_config_dict["MAIL_TYPE"]
        be_subject     = be_config_dict["SUBJECT"].format(process_code, now_string)
        process_name   = self.config["METADATA"]["PROCESS_NAME"]
        process_code   = self.config["METADATA"]["PROCESS_CODE"]
        # Set up email parameters
        attachments = [] # add logic if one or more BE needs attachments
        body_fields = [process_name, error] # add logic if one or more BE needs different body_fields
        # Set up business exception object, do not change
        self.be_info = {
            "BODY_FILE": be_body_file,
            "SUBJECT": be_subject,
            "MAIL_TYPE": be_mail_type,
            "ATTACHMENTS": attachments,
            "BODY_FIELDS": body_fields
        }

    def script_function_1(self):
        logging.info(f"--- DESCRIPTIVE START FOR SPRIPT 1 ---")
        # Run script after this
        logging.info(f"Reached example step 1 {self.template_parameter_1}")
        logging.info(f"DESCRIPTIVE END FOR SPRIPT 1")
        # raise BusinessException("This is a business exception")
        return 0

    def script_function_2(self):
        logging.info(f"--- DESCRIPTIVE START FOR SPRIPT 2 ---")
        # Run script after this
        #raise BusinessException("Testing business exceptions") # Testing business exception
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