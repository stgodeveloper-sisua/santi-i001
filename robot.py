# file_name:   robot.py
# created_on:  2023-03-17 ; vicente.diaz ; juanpablo.mena ; juan.montengero
# modified_on: 2023-04-04 ; guillermo.konig
# modified_on: 2023-05-16 ; vicente.diaz ; added send_monitoring_report() and robot_tray_active
# modified_on: 2023-06-15 ; vicente.diaz ; added env in monitoring report function
# modified_on: 2023-06-28 ; vicente.diaz ; added kill processes in finally state
# modified_on: 2023-06-29 ; vicente.diaz ; added monitoring class to upload execution logs

import os
import sys
import time
import socket
import logging
import traceback
import PySimpleGUI as sg
from modules._fmw.fmw_utils import *
from modules._fmw.monitoring import Monitoring
from modules._fmw.send_email_utils import send_email, df2html
from main import MainProcess


def process_init(config: dict):
        # create process dependencies file
    os.system(command="pip freeze > requirements.txt")  
    # clean process_data folder before starting the execution
    if not "DELETE_PROCESS_DATA_BEFORE_EXE" in config["FRAMEWORK"]:
        delete_process_data = True
    else:
        delete_process_data = config["FRAMEWORK"]["DELETE_PROCESS_DATA_BEFORE_EXE"]
    if delete_process_data:
        delete_folder(folder_path=config["FRAMEWORK"]["PROCESS_DATA"]) 
        create_folder(folder_path=config["FRAMEWORK"]["PROCESS_DATA"])
    # create output folder before starting the execution
    create_folder(folder_path=config["FRAMEWORK"]["OUTPUT"])
    return 0


def send_monitoring_report(config:dict, dti_start_exe=dt.datetime.now(), 
                           status=None, processed_states=None, total_states=None):
    logging.info(f"--- Sending monitoring report ---")
    config_monitoring = config["EMAIL"]["MONITORING_REPORT"]
    process_code      = config["METADATA"]["PROCESS_CODE"]
    environment       = config["METADATA"]["ENVIRONMENT"]
    now_str           = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dti_final_exe     = dt.datetime.now()
    file_exe_report   = os.path.join(config["FRAMEWORK"]["PROCESS_DATA"], f"{now_str}_exe_report.json")
    df_exe_report     = build_execution_summary_table(config, file_exe_report, dti_start_exe, dti_final_exe, 
                                                      status, processed_states, total_states)
    exe_log_file      = get_last_log_file(log_folder=config["FRAMEWORK"]["LOG_FOLDER"])
    send_email(to=config_monitoring["RECIPIENTS"]
                ,cc=[]
                ,bcc=[]
                ,subject=config_monitoring["SUBJECT"].format(process_code, now_str)
                ,body_file=config_monitoring["BODY_FILE"]
                ,body_fields=[process_code, df2html(df=df_exe_report)]
                ,wrapper_file=config["EMAIL"]["WRAPPER_FILE"]
                ,attachments=[exe_log_file, file_exe_report]
                ,environment=environment)
    return 0


if __name__ == "__main__":
    try:
        working_directory = os.path.dirname(__file__)
        os.chdir(working_directory)
        t1 = time.time()
        dti_start_time = dt.datetime.now()
        # Read execution arguments                
        config       = read_config()
        args         = sys.argv
        state_idx    = int(args[1]) if len(args) > 1 else 1
        # Create instance
        robot_instance = MainProcess(state_idx=state_idx)
        # Get config parameters
        process_name = config["METADATA"]["PROCESS_NAME"]
        n_area       = config["METADATA"]["N_AREA"]
        n_gerencia   = config["METADATA"]["N_GERENCIA"]
        environment  = config["METADATA"]["ENVIRONMENT"].upper()
        # Start execution logging instance
        start_logging(log_folder=config["FRAMEWORK"]["LOG_FOLDER"]
                    , label=config["METADATA"]["PROCESS_CODE"]
                    , logs_level="INFO"
                    , show_console_logs=True
                    , save_logs=True)
        logging.info(f"Hostname: {socket.gethostname()}")
        logging.info(f"Process name: {process_name}")
        logging.info(f"N_Area: {n_area}")
        logging.info(f"N_Gerencia: {n_gerencia}")
        logging.info(f"Environment: {environment}")
        # Build system tray notification
        robot_tray_active = config["FRAMEWORK"]["ACTIVE_ROBOT_TRAY"]
        logging.info(f"Robot tray active: {robot_tray_active}")
        if robot_tray_active:
            logging.info("Starting robot tray ...")
            tray_time   = (1000, 1000)
            quit_option = "Stop bot"
            menu_def    =  ["bot_menu", quit_option] 
            tray        = sg.SystemTray(menu=menu_def, data_base64=sg.EMOJI_BASE64_HAPPY_CONTENT)           
            tray.show_message('Starting process', f'Process: {process_name} | State: {state_idx}', time=tray_time)
        # Kill main process applications to ensure a clean execution
        if config["FRAMEWORK"]["KILL_PROCESSES"]:
            kill_processes(config["FRAMEWORK"]["LIST_PROCESESS_TO_KILL"])
         # Prepare process folders for execution
        process_init(config=config)
        # Start process thread
        robot_instance.start()
        # Monitoring thread
        while robot_instance.is_alive():
            if robot_tray_active:
                event = tray.read(timeout=20) # ms            
                if event in (sg.WINDOW_CLOSED, quit_option):    
                    # Stop execution if the robot window was closed   
                    if robot_instance.current_workflow != None:
                        robot_instance.status = "STOPPED"
                        tray_time             = (500, 500)
                    break
    except Exception as error:
        exception = traceback.format_exc()
        logging.error(f"Error found: {exception}")
    finally:
        # Ending execution
        if config["FRAMEWORK"]["KILL_PROCESSES"]:
            kill_processes(config["FRAMEWORK"]["LIST_PROCESESS_TO_KILL"])
        total_states         = get_total_states()
        robot_current_state  = max(robot_instance.state_idx, 1)
        success_state_offset = -1 if robot_instance.status ==  "SUCCESS" else 0
        robot_current_state  += success_state_offset
        # Update robot tray emoji
        if robot_tray_active:
            if robot_instance.status == "STOPPED":
                tray.update(data_base64=sg.EMOJI_BASE64_DEAD)
            elif robot_instance.status == "FAILED":
                tray.update(data_base64=sg.EMOJI_BASE64_FRUSTRATED)
            elif robot_instance.status == "WARNING":
                tray.update(data_base64=sg.EMOJI_BASE64_FRUSTRATED)
            elif robot_instance.status ==  "SUCCESS":
                tray.update(data_base64=sg.EMOJI_BASE64_HAPPY_BIG_SMILE)
        # Log execution outcome
        states_completed_message = f"States Completed: {robot_current_state}/{total_states}"
        final_execution_msg = f"Process: {process_name} | Status: {robot_instance.status.upper()}"
        logging.info(final_execution_msg)
        logging.info(states_completed_message)
        if robot_tray_active:
            logging.info("Closing robot tray ...")
            tray.show_message(title="Process finished", message=f"{final_execution_msg}", time=tray_time)
            tray.close()
    # Final script logs
    logging.info(f"Elapsed time (robot.py): {dt.timedelta(seconds=time.time()-t1)}")
    # Upload execution log to monitoring database
    if environment in ["QAS", "PRD"]:
        try:
            monitor = Monitoring(config=config)
            monitor.uplog()
        except Exception as error:
            logging.error(f"Could no upload execution logs using monitoring.py. Error: {error}")
    # Send monitoring report
    if environment == "PRD":
        send_monitoring_report(config=config, dti_start_exe=dti_start_time, status=robot_instance.status.upper(),
                            processed_states=robot_current_state, total_states=total_states)
    logging.info("--------- robot.py finished ---------")
