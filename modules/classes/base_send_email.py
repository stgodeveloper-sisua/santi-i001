# file_name:   base_send_email.py
# created_on:  2023-03-28 ; juan.montenegro
# modified_on: 2023-03-29 ; juan.montenegro

import logging
import os
import sys
import os
import zipfile
import shutil
from datetime import date

sys.path.append(os.path.abspath(
                    os.path.dirname(
                    os.path.dirname(
                    os.path.abspath(__file__)))))
                    
# from libraries.outlook_utils.outlook_utils import *
from _fmw.fmw_utils import *
from _fmw.send_email_utils import read_recipients_file, send_email


class SendEmailBase():
    def __init__(self, config:dict
                 , mail_type:str="EXECUTION_REPORT"
                 , sheet_recipients:str="base"
                 , compress_report:bool=False
                 , attach_files:bool=True
                 ):
        self.config                                = config
        self.environment                           = self.config["METADATA"]["ENVIRONMENT"].upper()
        self.process_code                          = self.config["METADATA"]["PROCESS_CODE"]
        self.process_name                          = self.config["METADATA"]["PROCESS_NAME"]                
        self.recipients_path                       = self.config["EMAIL"]["RECIPIENTS_PATH"]
        self.wrapper_file                          = self.config["EMAIL"]["WRAPPER_FILE"]
        self.mail_type                             = mail_type
        self.config_email                          = self.config["EMAIL"][mail_type]   
        self.compress_report                       = compress_report   
        self.sheet_recipients                       = sheet_recipients
        self.attach_files                          = attach_files
        self.body_file                             = None
        self.body_fields                           = [self.process_name]
        self.list_source_files                     = []
        self.list_target_files                     = []
        self.attachments                           = []
        self.size_zip                              = 0
        self.size_all_files                        = 0
        self.zip_fullname                          = None
        self.subject                               = ""
        self.cfg_path_mail_body_file               = None
        self.cfg_size_outlook__attachments         = 100
        self.cfg_zip_path                          = None          
        self.cfg_zip_filename                      = None
        if "SIZE_LIMIT_ATTACHMENT" in self.config_email:
            self.cfg_size_outlook__attachments         = self.config["EMAIL"]["SIZE_LIMIT_ATTACHMENT"]
        if "SUBJECT" in self.config_email:
            self.subject                               = self.config_email["SUBJECT"]
        if "BODY_FILE" in self.config_email:    
            self.cfg_path_mail_body_file               = self.config_email["BODY_FILE"]
        if "BODY_FILE_SHAREPOINT" in self.config_email:
            self.cfg_path_mail_body_file_sharepoint    = self.config_email["BODY_FILE_SHAREPOINT"]
        if "ZIP_REPORT_PATH" in self.config_email:
            self.cfg_zip_path                          = self.config_email["ZIP_REPORT_PATH"]            
            self.cfg_zip_filename                      = self.config_email["ZIP_REPORT_FILENAME"]                  
    
    def __del__(self):
        pass

    def set_zip_fullname(self):
        if self.cfg_zip_path != None:
            self.zip_fullname = os.path.join(self.cfg_zip_path)
        if self.cfg_zip_filename != None:
            self.zip_fullname = os.path.join(self.zip_fullname, self.cfg_zip_filename)

    def read_recipients(self):
        if self.sheet_recipients != None:     
            self.to,self.cc,self.bcc = read_recipients_file(recipients_file_path=self.recipients_path
                                                                ,environment=self.environment 
                                                                ,mail_type=self.mail_type.upper()
                                                                ,sheet=self.sheet_recipients)


    def zip_report(self):
       if (self.zip_fullname != None):
        with zipfile.ZipFile(self.zip_fullname, "w") as zip_report_file:
            for file_source in self.list_source_files :
                logging.info(f"Zipping file {file_source} in {self.zip_fullname}")  
                zip_report_file.write(file_source
                                , arcname=os.path.basename(file_source)
                                , compress_type=zipfile.ZIP_LZMA
                                , compresslevel = 9) 

    def copy_files(self):
        for file_source, file_target in zip(self.list_source_files , self.list_target_files ):
            if not file_target is None:
                logging.info(f"Copying file {file_source} to file {file_target}")  
                shutil.copy(file_source, file_target) 

    def set_subject(self):
        self.subject = self.subject.format(date.today().strftime("%d-%m-%y"))

    def set_attachments(self):
        self.size_zip = 0
        if (self.zip_fullname != None):
            self.size_zip = int(get_size(self.zip_fullname, 'mb'))
        self.size_all_files = 0
        self.attachments = []
        for file_source in self.list_source_files :
            self.size_all_files += int(get_size(file_source, 'mb'))
        if (self.size_all_files < self.cfg_size_outlook__attachments):            
            for file_source, file_target in zip(self.list_source_files , self.list_target_files ):
                if file_target is None:
                    self.attachments.append(os.path.abspath(file_source))
                else:
                    self.attachments.append(os.path.abspath(file_target))
        elif self.compress_report and (self.size_zip < self.cfg_size_outlook__attachments):
            self.attachments = [self.zip_fullname]

    def set_body_file(self):
        self.body_file = self.cfg_path_mail_body_file
        if (self.size_all_files > self.cfg_size_outlook__attachments):
            self.body_file = self.cfg_path_mail_body_file_sharepoint

    def set_body_fields(self):
        self.body_fields = [self.process_name]

    def send_email(self):
        self.read_recipients()
        self.set_zip_fullname()
        if (self.compress_report):
            self.zip_report()   
        self.set_subject()         
        self.copy_files()
        if self.attach_files:
            self.set_attachments()
        self.set_body_file()
        self.set_body_fields()
        send_email(to=self.to
                ,cc=self.cc
                ,bcc=self.bcc
                ,subject=self.subject
                ,body_file=self.body_file
                ,body_fields=self.body_fields
                ,wrapper_file=self.wrapper_file
                ,attachments=self.attachments)


if __name__ == "__main__":
    start_logging(logs_level="DEBUG", show_console_logs=True, save_logs=False)
    logging.info(f"--------- Script {os.path.basename(__file__)} started ---------")
    config = read_config()  
    state = SendEmailBase(config=config)
    state.send_email()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")    