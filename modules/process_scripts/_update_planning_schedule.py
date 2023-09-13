# file_name:   _update_planning_schedule.py
# created_on:  2023-09-09 ; santiago.garcia ; updated self.state_name attribute
# modified_on: 2023-09-10 ; santiago.garcia

# DESCRIPTION OF THE PROCESS:
#   This RPA process reads the input file used by a Solution Consultant to update the hours registered for each one of the unfinnished projects assigned to him/her in ClickUp.
#   The process will contemplate 5 independant states:
#       1. Generate the worktray: This state will read the input file and generate the worktray file of the process.
#       2. Extract the ClickUp updated data: For each row of the worktray, the robot will extract all the project 
#          data from ClickUp with help of the API of the platform. Then, it will save this data into a .csv file.
#       3. Process the Clickup data: For each row in the worktray, this state will read each '.csv' file and use that data to update the scheduling files.
#          Then, it will save the updated scheduling files into the 'process_data' folder.
#       4. Send the output file to the Solution Consultant One Drive and update the input (memory) file:
#          This state will send the updated scedules files into the Solution Consultant folder and update the input file with the new data.
#       5. Send the RPA execution report: This state will send an email to the Solution Consultant with the summary of the execution of this RPA.

import os
import sys
import logging
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # move to modules 
from _fmw.fmw_utils import *
from _fmw.fmw_classes import BusinessException
from classes.robot_date import RobotDate
from process_scripts._base_process_class import ProcessBase
import generate_worktray as generator

class UpdatePlanningSchedule(ProcessBase, RobotDate):
    def __init__(self, config:dict):
        ProcessBase.__init__(self, config=config)
        RobotDate.__init__(self ,config=config) 
        self.state_name = self.state_name = type(self).__name__ # Get the class name, do not change     
        # workflow parameters
        # STATE 1 
        self.process_input_file           = config["GLOBAL"]["PROCESS_INPUT_FILE"]
        self.process_input_file_name      = os.path.basename(self.process_input_file)
        self.process_input_file_worksheet = config["GLOBAL"]["PROCESS_INPUT_FILE_WORKSHEET"]
        self.process_input_file_columns   = config["GLOBAL"]["PROCESS_INPUT_FILE_COLUMNS"]
        self.worktray_template            = config["GLOBAL"]["WORKTRAY_FILE_TEMPLATE"]
        self.worktray_file                = config["GLOBAL"]["WORKTRAY_FILE"]
        self.worktray_base_worksheet      = config["GLOBAL"]["WORKTRAY_BASE_WORKSHEET"]
        self.process_data_files_worksheet = config["GLOBAL"]["PROCESS_DATA_FILES_WORKSHEET"]
        self.planning_file_path           = config["GLOBAL"]["PLANNING_FILE_PATH"]
        self.raw_file_path                = config["GLOBAL"]["RAW_FILE_PATH"]
        self.clickup_data_path            = config["GLOBAL"]["CLICKUP_DATA_PATH"]
        self.processed_file_path          = config["GLOBAL"]["PROCESSED_FILE_PATH"]
        self.raw_file_path                = config["GLOBAL"]["RAW_FILE_PATH"]
        self.worktray_filter_column       = config["GLOBAL"]["WORKTRAY_FILTER_COLUMN"]
        self.process_data_file_timestamp  = config["GLOBAL"]["PROCESS_DATA_TIMESTAMP_FORMAT"]
        #This datetime is in the .txt file inside the process_data folder
        self.processing_datetime          = self.execution_date
        # Business and System exceptions variables
        self.business_exception_one       = config["GLOBAL"]["BUSINESS_EXCEPTION_ONE"]
        self.business_exception_two       = config["GLOBAL"]["BUSINESS_EXCEPTION_TWO"]
        self.business_exception_three     = config["GLOBAL"]["BUSINESS_EXCEPTION_THREE"]
        # STATE 2
        self.clickup_api_key              = self.config_env["CLICKUP_API_KEY"]
        
        # These variables already exist by inheritance
        # self.config              = config
        # self.environment         = config["METADATA"]["ENVIRONMENT"].upper()
        # self.config_env          = config[self.environment]
        # config                   = config["GLOBAL"]
        # self.config_fmw          = config["FRAMEWORK"]        
        # self.process_data        = self.config_fmw["PROCESS_DATA"]
        # self.config_emails       = self.config["EMAIL"]
        # self.now                 = dt.datetime.now()
        
    #CHECKPOINT
    def get_clickup_data(self):
        logging.info(f"--- CONNECTING INTO THE CLICKUP API ---")
        # Process Global variables
        CLICKUP_API_KEY = self.clickup_api_key
        #Now, we will connect into the Clickup API
        client = ClickUpClient(CLICKUP_API_KEY)
        # Example request | Creating a task in a list
        c = client.ClickUpClient(CLICKUP_API_KEY)
        # Finally, we will save the dataframe into the worktray file
        save_excel_file(worktray, file_path=WORKTRAY_FILE, sheet_name=WORKTRAY_BASE_SHEET)
        #After the saving, the headers of the excel file change the color format, so we will change it back to the original format
        reset_worktray_headers_format(WORKTRAY_FILE, WORKTRAY_BASE_SHEET, WORKTRAY_TEMPLATE)
        logging.info(f"Reached example step 1 {PROCESS_INPUT_FILE}")
        logging.info("CLICKUP PROJECTS DATA EXTRACTION FINISHED")
        return 0

    def run_workflow(self):
        logging.info(f"----- Starting state: {self.state_name} -----")
        try: # Add workflow in try block bellow
            # We will call the "generate_worktray" function from the "generate_worktray.py" file
            generator.generate_worktray(self) #State 1 
            self.get_clickup_data() #State 2   
            #self.script_function_2()
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
    process = UpdatePlanningSchedule(config=config)    
    process.run_workflow()
    # state.script_function_1()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     