# file_name:   process_utils.py
# created_on:  2023-03-28 ; sisua.developer
# modified_on: 2023-03-29 ; sisua.developer

import os
import sys
import logging
import datetime as dt
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))  # move to modules 
from _fmw.fmw_utils import *


def util_funtion_1():
    pass


if __name__ == "__main__":
    start_logging(logs_level="DEBUG", show_console_logs=True, save_logs=False)
    logging.info(f"--------- Script {os.path.basename(__file__)} started ---------") 
    util_funtion_1()
    logging.info(f"--------- Script {os.path.basename(__file__)} finished ---------")     