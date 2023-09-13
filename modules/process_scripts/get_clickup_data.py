# file_name:   extract_clickup_data.py
# created_on:  2023-09-13 ; santiago.garcia
# modified_on: 2023-09-13 ; santiago.garcia

#----------------------------------------------------DESCRIPTION OF THE FUNCTION -------------------------------------------------#
# The first step of this module is to connect to the Clickup API. We will use our access token saved in the config.jsonc file.
# One we connected into the API, for each project code, we will find a match of the project code versus the total proyect list name.
# If a match it's found, then we will extract all the data of the project (tasks, time estimated, time tracked, etc) and we will
# save it into a file of the process_data folder. If there's no match, we will add an observation, and keep the FALSE boolean in
# the worktray ceLL.
#---------------------------------------------------------------------------------------------------------------------------------#

import os
import sys
import logging
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # move to modules 
from _fmw.fmw_utils import *
from _fmw.fmw_classes import BusinessException
from clickupython import client

def get_clickup_data(self):
    logging.info(f"--- CONNECTING INTO THE CLICKUP API ---")
    # Process Global variables
    self.clickup_api_key = CLICKUP_API_KEY
    #Now, we will connect into the Clickup API
    client = ClickUpClient(API_KEY)
    # Example request | Creating a task in a list
    c = client.ClickUpClient(API_KEY)
    if task:
        print(task.id)
    # Finally, we will save the dataframe into the worktray file
    save_excel_file(worktray, file_path=WORKTRAY_FILE, sheet_name=WORKTRAY_BASE_SHEET)
    #After the saving, the headers of the excel file change the color format, so we will change it back to the original format
    reset_worktray_headers_format(WORKTRAY_FILE, WORKTRAY_BASE_SHEET, WORKTRAY_TEMPLATE)
    logging.info(f"Reached example step 1 {PROCESS_INPUT_FILE}")
    logging.info("CLICKUP PROJECTS DATA EXTRACTION FINISHED")
    return 0