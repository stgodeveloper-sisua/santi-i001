# file_name:   generate_worktray.py
# created_on:  2023-09-10 ; santiago.garcia
# modified_on: 2023-09-10 ; santiago.garcia

#----------------------------------------------------DESCRIPTION OF THE STATE 1 -----------------------------------------------#
# This state will read the input file of the process (PROCESS_INPUT_FILE) and generate the worktray file of the process.
# The worktray file is an excel file that contains all the information that the robot needs to process the data (located in Sharepoint).
# Once the input (memory) file is read, the robot will keep only the projects that aren't finished yet.
# In other words, the robot will keep only the rows that have the "FALSE" value in the "project_finished" column.
# Then, the robot will generate the timestamp for the process data file following the format that 
#   we defined in the config file (PROCESS_DATA_FILE_TIMESTAMP_FORMAT).
# Finally, the robot will save the worktray file into the worktray file path (WORKTRAY_FILE).
#------------------------------------------------------------------------------------------------------------------------------#

import os
import sys
import logging
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # move to modules 
from _fmw.fmw_utils import *
from _fmw.fmw_classes import BusinessException

def generate_worktray(self):
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
        PLANNING_FILE_PATH                 = self.planning_file_path
        RAW_FILE_PATH                      = self.raw_file_path
        PROCESSED_FILE_PATH                = self.processed_file_path
        CLICKUP_DATA_PATH                  = self.clickup_data_path
        RAW_FILE_NAME                      = os.path.basename(PLANNING_FILE_PATH)
        PROCESSED_FILE_NAME                = os.path.basename(PROCESSED_FILE_PATH)
        CLICKUP_DATA_NAME                  = os.path.basename(CLICKUP_DATA_PATH)
        PROCESSED_FILE_PATH                = self.processed_file_path
        WORKTRAY_FILTER_COLUMN             = self.worktray_filter_column
        PROCESS_DATA_FILE_TIMESTAMP_FORMAT = self.process_data_file_timestamp
        PROCESSING_DATETIME                = self.processing_datetime
        # Business and System exceptions variables
        BUSINESS_EXCEPTION_ONE = self.business_exception_one # The input file wasn't found in the correct directory
        BUSINESS_EXCEPTION_TWO = self.business_exception_two # The input file has a different number of columns than the expected
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
            # We find and BusinessException here. If the user somehow remove at least one column from the input file, then the process must stop
            if total_columns != PROCESS_INPUT_FILE_COLUMNS:
                raise BusinessException(f"{BUSINESS_EXCEPTION_TWO}".format(total_columns, PROCESS_INPUT_FILE_COLUMNS))
            insert_data = [] # We will insert the data into the worktray row by row, so we will create a list for each row
            #We will create the timestamp for the process data file following the format that we defined in the config file
            timestamp              = PROCESSING_DAY.strftime(PROCESS_DATA_FILE_TIMESTAMP_FORMAT)
            project_id             = worktray.iloc[row_idx, 0]
            project_finished       = worktray.iloc[row_idx, 1]
            planning_file_path     = worktray.iloc[row_idx, 2]
            raw_file_path          = "{}".format(RAW_FILE_PATH).format(timestamp, project_id, RAW_FILE_NAME)
            clickup_data_extracted = worktray.iloc[row_idx, 4]
            clickup_data_path      = "{}".format(CLICKUP_DATA_PATH).format(timestamp, project_id, CLICKUP_DATA_NAME)
            processed_file_path    = "{}".format(PROCESSED_FILE_PATH).format(timestamp, project_id, PROCESSED_FILE_NAME)
            planning_file_updated  = worktray.iloc[row_idx, 7]
            execution_datetime     = PROCESSING_DATETIME
            observations = None # We may add some observations in a late stage of the process, so we will leave this column empty for now
            
            #Now we insert this list into the worktray row
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