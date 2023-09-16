# file_name:   _update_planning_schedule.py
# created_on:  2023-09-09 ; santiago.garcia ; updated self.state_name attribute
# modified_on: 2023-09-10 ; santiago.garcia

import os
import sys
import logging
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # move to modules 
from _fmw.fmw_utils import *
from _fmw.fmw_classes import BusinessException
from classes.robot_date import RobotDate
from process_scripts._base_process_class import ProcessBase
import requests
import json

class WeeklyPlanProposal(ProcessBase, RobotDate):
    def __init__(self, config:dict):
        ProcessBase.__init__(self, config=config)
        RobotDate.__init__(self ,config=config)
        # These variables already exist by inheritance
        # self.config              = config
        # self.environment         = config["METADATA"]["ENVIRONMENT"].upper()
        # self.config_env          = config[self.environment]
        # config                   = config["GLOBAL"]
        # self.config_fmw          = config["FRAMEWORK"]        
        # self.process_data        = self.config_fmw["PROCESS_DATA"]
        # self.config_emails       = self.config["EMAIL"]
        # self.now                 = dt.datetime.now()
        #-------------------------------------------- WORKFLOW PARAMETERS -------------------------------------------------#
        self.state_name = self.state_name = type(self).__name__ # Get the class name, do not change  
        #------------------------------------------------------------------------------------------------------------------#
        #                                             STATE 1 PARAMETERS                                                   #
        #------------------------------------------------------------------------------------------------------------------#
        self.process_input_file             = config["GLOBAL"]["PROCESS_INPUT_FILE"]
        self.process_input_file_name        = os.path.basename(self.process_input_file)
        self.process_input_file_worksheet   = config["GLOBAL"]["PROCESS_INPUT_FILE_WORKSHEET"]
        self.process_input_file_columns     = config["GLOBAL"]["PROCESS_INPUT_FILE_COLUMNS"]
        self.worktray_template              = config["GLOBAL"]["WORKTRAY_FILE_TEMPLATE"]
        self.worktray_file                  = config["GLOBAL"]["WORKTRAY_FILE"]
        self.worktray_base_worksheet        = config["GLOBAL"]["WORKTRAY_BASE_WORKSHEET"]
        self.process_data_files_worksheet   = config["GLOBAL"]["PROCESS_DATA_FILES_WORKSHEET"]
        self.planning_file_path             = config["GLOBAL"]["PLANNING_FILE_PATH"]
        self.raw_file_path                  = config["GLOBAL"]["RAW_FILE_PATH"]
        self.clickup_data_path              = config["GLOBAL"]["CLICKUP_DATA_PATH"]
        self.processed_file_path            = config["GLOBAL"]["PROCESSED_FILE_PATH"]
        self.raw_file_path                  = config["GLOBAL"]["RAW_FILE_PATH"]
        self.worktray_filter_column         = config["GLOBAL"]["WORKTRAY_FILTER_COLUMN"]
        self.process_data_file_timestamp    = config["GLOBAL"]["PROCESS_DATA_TIMESTAMP_FORMAT"]
        #This datetime is in the .txt file inside the process_data folder
        self.processing_datetime            = self.execution_date
        # Business and System exceptions variables
        self.business_exception_one         = config["GLOBAL"]["BUSINESS_EXCEPTION_ONE"]
        self.business_exception_two         = config["GLOBAL"]["BUSINESS_EXCEPTION_TWO"]
        self.business_exception_three       = config["GLOBAL"]["BUSINESS_EXCEPTION_THREE"]
        self.business_exception_four        = config["GLOBAL"]["BUSINESS_EXCEPTION_FOUR"]
        #------------------------------------------------------------------------------------------------------------------#
        #                                             STATE 2 PARAMETERS                                                   #
        #------------------------------------------------------------------------------------------------------------------#
        # To get more details about the use of the Clickup API, please visit this link: https://clickup.com/api/
        self.clickup_api_key                = "pk_43148791_9XH0VJZAW94J0TD3MYQZZ1VPM8VUKEJO"
        self.clickup_space_id               = config["GLOBAL"]["CLICKUP_SPACE_ID"] # Default value: 10869625 -> Sisua Chile
        self.json_clickup_folders_data_path = config["GLOBAL"]["JSON_CLICKUP_FOLDERS_DATA_PATH"]
        self.clickup_api_folders_url        = "https://api.clickup.com/api/v2/space/{}/folder" # {} -> clickup_space_id

#----------------------------------------------------STATE 1 FUNCTION ------------------------------------------------------------------#
# This state will read the input file of the process (PROCESS_INPUT_FILE) and generate the worktray file of the process.
# The worktray file is an excel file that contains all the information that the robot needs to process the data (located in Sharepoint).
# Once the input (memory) file is read, the robot will keep only the projects that aren't finished yet.
# In other words, the robot will keep only the rows that have the "FALSE" value in the "project_finished" column.
# Then, the robot will generate the timestamp for the process data file following the format that 
#   we defined in the config file (PROCESS_DATA_FILE_TIMESTAMP_FORMAT).
# Finally, the robot will save the worktray file into the worktray file path (WORKTRAY_FILE).
#---------------------------------------------------------------------------------------------------------------------------------------#

    def _generate_worktray(self):
            logging.info(f"--- GENERATING THE WORKTRAY ---")
            # Process Global variables
            PROCESSING_DAY                     = self.processing_datetime
            PROCESS_INPUT_FILE                 = self.process_input_file
            PROCESS_INPUT_FILE_DIRECTORY       = os.path.dirname(PROCESS_INPUT_FILE)
            PROCESS_INPUT_FILE_WORKSHEET       = self.process_input_file_worksheet
            PROCESS_INPUT_FILE_COLUMNS         = self.process_input_file_columns
            WORKTRAY_TEMPLATE                  = self.worktray_template
            WORKTRAY_FILE                      = self.worktray_file
            WORKTRAY_BASE_SHEET                = self.worktray_base_worksheet
            PROCESS_INPUT_FILE_NAME            = self.process_input_file_name
            RAW_FILE_PATH                      = self.raw_file_path
            PROCESSED_FILE_PATH                = self.processed_file_path
            CLICKUP_DATA_PATH                  = self.clickup_data_path
            RAW_FILE_NAME                      = os.path.splitext(os.path.basename(RAW_FILE_PATH))[0]
            PROCESSED_FILE_NAME                = os.path.splitext(os.path.basename(PROCESSED_FILE_PATH))[0]
            CLICKUP_DATA_NAME                  = os.path.splitext(os.path.basename(CLICKUP_DATA_PATH))[0]
            PROCESSED_FILE_PATH                = self.processed_file_path
            WORKTRAY_FILTER_COLUMN             = self.worktray_filter_column
            PROCESS_DATA_FILE_TIMESTAMP_FORMAT = self.process_data_file_timestamp
            PROCESSING_DATETIME                = self.processing_datetime
            # Business and System exceptions variables
            BUSINESS_EXCEPTION_ONE = self.business_exception_one # The input file wasn't found in the correct directory
            BUSINESS_EXCEPTION_TWO = self.business_exception_two # Some input rows doesn't contains minimal inputs.
            BUSINESS_EXCEPTION_THREE = self.business_exception_three # The input file has a different number of columns than the expected
            logging.info(f"Robot processing day: {PROCESSING_DAY}")
            logging.info(f"Input file name: {PROCESS_INPUT_FILE_NAME}")
            # We want to make shure that this file exist. Otherwise, the process can't continue.
            if not os.path.exists(PROCESS_INPUT_FILE):
                raise BusinessException(f"{BUSINESS_EXCEPTION_ONE}".format(PROCESS_INPUT_FILE, PROCESS_INPUT_FILE_DIRECTORY))
            # If already exists an older worktray, then we will overwrite it
            if os.path.exists(WORKTRAY_FILE):
                os.remove(WORKTRAY_FILE)
            shutil.copy(WORKTRAY_TEMPLATE, WORKTRAY_FILE) # We copy the template into the worktray file
            #We will read the excel input file and convert it into a dataframe, considering the first row as the header
            df = pd.read_excel(PROCESS_INPUT_FILE, sheet_name=PROCESS_INPUT_FILE_WORKSHEET, header=0)
            #We keep only the rows that we need to process (the ones that have the "FALSE" value in the filter column)
            worktray = df[df[WORKTRAY_FILTER_COLUMN] == False]
            total_rows = len(worktray)
            total_columns = len(worktray.columns)
            logging.info(f"Total rows to process: {total_rows}")
            logging.info(f"Total columns in worktray: {total_columns}")
            for row_idx in range(total_rows):
                # Business Exception: If one row doesn't contains at least the four principal values, then we add an observation and continue with the next row.
                if worktray.iloc[row_idx, 0] == None or worktray.iloc[row_idx, 1] == None or worktray.iloc[row_idx, 2] == None or worktray.iloc[row_idx, 3] == None:
                    worktray.iloc[row_idx, 10] = BUSINESS_EXCEPTION_TWO
                    continue
                # We find and BusinessException here. If the user somehow remove at least one column from the input file, then the process must stop
                if total_columns != PROCESS_INPUT_FILE_COLUMNS:
                    raise BusinessException(f"{BUSINESS_EXCEPTION_THREE}".format(total_columns, PROCESS_INPUT_FILE_COLUMNS))
                insert_data = [] # We will insert the data into the worktray row by row, so we will create a list for each row
                #We will create the timestamp for the process data file following the format that we defined in the config file
                timestamp              = PROCESSING_DAY.strftime(PROCESS_DATA_FILE_TIMESTAMP_FORMAT)
                customer               = worktray.iloc[row_idx, 0]
                project_id             = worktray.iloc[row_idx, 1]
                project_finished       = worktray.iloc[row_idx, 2]
                planning_file_path     = worktray.iloc[row_idx, 3]
                raw_file_path          = "{}".format(RAW_FILE_PATH).format(timestamp, project_id, RAW_FILE_NAME)
                clickup_data_extracted = worktray.iloc[row_idx, 5]
                clickup_data_path      = "{}".format(CLICKUP_DATA_PATH).format(timestamp, project_id, CLICKUP_DATA_NAME)
                processed_file_path    = "{}".format(PROCESSED_FILE_PATH).format(timestamp, project_id, PROCESSED_FILE_NAME)
                planning_file_updated  = worktray.iloc[row_idx, 8]
                execution_datetime     = PROCESSING_DATETIME
                observations = None # We may add some observations in a late stage of the process, so we will leave this column empty for now
                
                #Now we insert this list into the worktray row
                insert_data.append(customer)
                insert_data.append(project_id)
                insert_data.append(project_finished)
                insert_data.append(planning_file_path)
                insert_data.append(raw_file_path)
                insert_data.append(clickup_data_extracted)
                insert_data.append(clickup_data_path)
                insert_data.append(processed_file_path)
                insert_data.append(planning_file_updated)
                insert_data.append(execution_datetime)
                insert_data.append(observations)
                #After we add all the data into this list, we will insert it into the dataframe row by row
                worktray.loc[row_idx] = insert_data
            # Finally, we will save the dataframe into the worktray file
            save_excel_file(worktray, file_path=WORKTRAY_FILE, sheet_name=WORKTRAY_BASE_SHEET)
            #After the saving, the headers of the excel file change the color format, so we will change it back to the original format
            reset_worktray_headers_format(WORKTRAY_FILE, WORKTRAY_BASE_SHEET, WORKTRAY_TEMPLATE)
            logging.info(f"Reached example step 1 {PROCESS_INPUT_FILE}")
            logging.info("WORKTRAY_FILE GENERATED SUCCESFULLY")
            return 0 
#----------------------------------------------------STATE 2 FUNCTION ------------------------------------------------------------------#
# This state will connect to the Clickup API and extract the data of the projects that we need to process.
#---------------------------------------------------------------------------------------------------------------------------------------#

    def get_clickup_data(self): #State 2 function
        # To get more details about the use of the Clickup API, please visit this link: https://clickup.com/api/
        # Process Global variables
        CLICKUP_API_KEY = self.clickup_api_key
        CLICKUP_API_URL     = self.clickup_api_folders_url .format(self.clickup_space_id)
        WORKTRAY_FILE = self.worktray_file
        WORKTRAY_BASE_WORKSHEET = self.worktray_base_worksheet
        JSON_CLICKUP_FOLDERS_DATA_PATH = self.json_clickup_folders_data_path
        CLICKUP_SPACE_ID = self.clickup_space_id
        CLICKUP_API_URL = f"{CLICKUP_API_URL}".format(CLICKUP_SPACE_ID)
        # Business and System exceptions variables
        BUSINESS_EXCEPTION_FOUR = self.business_exception_four # Problems connecting to the Clickup API

        logging.info(f"--- EXTRACTING THE SISUACL CLICKUP FOLDERS BY API ---")
        query = {
        "archived": "false"
        }
        headers = {"Authorization": CLICKUP_API_KEY}
        try:
            response = requests.get(CLICKUP_API_URL, headers=headers, params=query)
        except Exception as error:
            logging.error(f"--- ERROR CONNECTING INTO THE CLICKUP API ---")
            logging.error(f"{BUSINESS_EXCEPTION_FOUR}")
            raise BusinessException(BUSINESS_EXCEPTION_FOUR)
        logging.info(f"--- EXTRACTION SUCCESFUL ---")

        # We will save this JSON data into the process_data_folder into a json extension
        logging.info("SAVING JSON CLICKUP DATA INTO THE PROCESS DATA FOLDER")
        clickup_folders_data = response.json()
        with open(JSON_CLICKUP_FOLDERS_DATA_PATH, 'w') as f:
            json.dump(clickup_folders_data, f)
        # Reading the worktray
        worktray = pd.read_excel(WORKTRAY_FILE, sheet_name=WORKTRAY_BASE_WORKSHEET)
        # For each folder in the response, we will extract the data that we need and save it into the worktray file
        for folder in response.json()["folders"]:
            folder_id    = folder["id"]
            folder_name  = folder["name"]
            folder_lists = folder["lists"]
            # We will iterate over the rows of the worktray file
            for row_idx in range(len(worktray)):
                # Getting the project code
                project_id = worktray.iloc[row_idx, 1]
                # If the project code is equal to the folder name, then we will save the folder data into the worktray row
                if project_id == folder_name:
                    worktray.iloc[row_idx, 3] = folder_id
                    worktray.iloc[row_idx, 5] = folder_lists
                    break


        # Now, we will read the worktra file and convert it into a dataframe
        df = pd.read_excel(WORKTRAY_FILE, sheet_name=WORKTRAY_BASE_WORKSHEET)
        #CHECK POINT: JSON DATA EXTRACTED IN 50% OF DEVELOPMENT
        # We will iterate over the rows of the dataframe
        for row_idx in range(len(df)):
            #Getting the project code
            project_id = df.iloc[row_idx, 0]
            # Now, with the "response" request that we did from the API we will extract the
            # project that contains the code retrieved from this worktray row
            for project in response.json()["teams"][0]["projects"]:
                if project["name"] == project_id:
                    # We will save the project data into the worktray row
                    df.iloc[row_idx, 4] = project
                    break
        logging.info("CLICKUP PROJECTS DATA EXTRACTION FINISHED")
        return 0

    def run_workflow(self):
        logging.info(f"----- Starting state: {self.state_name} -----")
        try: # Add workflow in try block bellow
            self._generate_worktray() #State 1
            self.get_clickup_data() #State 2
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
    process = WeeklyPlanProposal(config=config)    
    process.run_workflow()
    # state.script_function_1()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     