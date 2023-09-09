# file_name:   send_exceptions_emails.py
# created_on:  2023-03-28 ; guillermo.konig
# modified_on: 2023-03-29 ; guillermo.konig

import os
import sys
import logging
import datetime as dt  
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))           
from _fmw.fmw_utils import *
from _fmw.send_email_utils import read_recipients_file, send_email, df2html


class ExceptionEmails:
    def __init__(self, config:dict):
        self._state_name    = "_TEMPLATE_STATE_NAME"  # Change with state name      
        self._config        = config
        self._environment   = self._config["METADATA"]["ENVIRONMENT"].upper()
        self._config_env    = self._config[self._environment]
        self._config_global = self._config["GLOBAL"]        
        self._state_success = True
        self._robot_stop    = False
        # Set up now date
        self._now_string = dt.datetime.now().strftime("%d-%m-%y")
        # Get config parameters
        self._config_emails       = self._config["EMAIL"]
        self._process_name        = self._config["METADATA"]["PROCESS_NAME"]
        self._process_code        = self._config["METADATA"]["PROCESS_CODE"]
        ## Get general email parameters 
        self._email_wrapper       = self._config_emails["WRAPPER_FILE"]
        self._recipients_file     = self._config_emails["RECIPIENTS_PATH"]
        ## Get system exception parameters
        self.sys_exc_body_file    = self._config_emails["SYS_EXC_REPORT"]["BODY_FILE"]
        self.sys_exc_email_type   = "SYS_EXC_REPORT"
        self.sys_exc_subject      = self._config_emails["SYS_EXC_REPORT"]["SUBJECT"].format(self._process_code, self._now_string)
        ## Get user system exception parameters 
        self._send_user_exception = self._config_emails["ENABLE_CLIENT_SYS_EXC"]
        self.usr_exc_body_file    = self._config_emails["SYS_EXC_REPORT_USER"]["BODY_FILE"]
        self.usr_exc_email_type   = "SYS_EXC_REPORT_USER"
        self.usr_exc_subject      = self._config_emails["SYS_EXC_REPORT_USER"]["SUBJECT"].format(self._process_code, self._now_string)

    def send_user_system_exception(self):
        logging.info("----- Sending user system exception -----")
        # Check if send user email
        if self._send_user_exception:
            # Set up outlook parameters
            attachments = []
            body_fields = [self._process_name]
            to,cc,bcc   = read_recipients_file(recipients_file_path = self._recipients_file,
                                            environment = self._environment,
                                            mail_type = self.usr_exc_email_type,
                                            sheet = "base")
            # Send email
            send_email(to=to,cc=cc,bcc=bcc,
                    environment=self._environment,
                    subject=self.usr_exc_subject,
                    body_file=self.usr_exc_body_file,
                    body_fields=body_fields,
                    wrapper_file=self._email_wrapper,
                    attachments=attachments)
            return True
        else:
            logging.warning("Send user exception is disabled by config")
    
    def send_system_exception(self, exception_df:str, exc_source:str, exc_message:str):
        logging.info("----- Sending system exception -----")
        # Tarnsform traceback to html
        exception_html = df2html(df=exception_df)
        # Set up outlook parameters
        attachments = []
        body_fields = [self._process_name, exc_message, exc_source, exception_html]
        to,cc,bcc   = read_recipients_file(recipients_file_path = self._recipients_file,
                                           environment = self._environment,
                                           mail_type = self.sys_exc_email_type,
                                           sheet = "base")
        # Send email
        send_email(to=to,cc=cc,bcc=bcc,
                   environment=self._environment,
                   subject=self.sys_exc_subject,
                   body_file=self.sys_exc_body_file,
                   body_fields=body_fields,
                   wrapper_file=self._email_wrapper,
                   attachments=attachments)
        return True

    def send_business_exception(self, be_info:dict):
        logging.info("----- Sending business exception -----")
        # Get BE info
        be_body_file =  be_info["BODY_FILE"]
        be_subject   =  be_info["SUBJECT"] 
        be_mail_type =  be_info["MAIL_TYPE"]
        attachments  =  be_info["ATTACHMENTS"]
        body_fields  =  be_info["BODY_FIELDS"]
        # Set up recipients
        to,cc,bcc   = read_recipients_file(recipients_file_path = self._recipients_file,
                                        environment = self._environment,
                                        mail_type = be_mail_type,
                                        sheet = "base")
        if "TO" in be_info.keys():
            to = to + be_info["TO"]
        # Send email
        send_email(to=to,cc=cc,bcc=bcc,
                environment=self._environment,
                subject=be_subject,
                body_file=be_body_file,
                body_fields=body_fields,
                wrapper_file=self._email_wrapper,
                attachments=attachments)
        return True

if __name__ == "__main__":
    print (os.path.dirname(__file__))
    config = read_config()
    start_logging(log_folder=config["FRAMEWORK"]["LOG_FOLDER"], label=config["METADATA"]["PROCESS_CODE"], logs_level="INFO", show_console_logs=True, save_logs=True)
    logging.info(f"--------- Script {os.path.basename(__file__)} started ---------")    
    state = ExceptionEmails(config=config)
    # state.run_state()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     