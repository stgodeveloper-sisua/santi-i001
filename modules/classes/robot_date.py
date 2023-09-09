# file_name:   robot_date.py
# created_on:  2023-03-14 ; joaquín.sandoval
# modified_on: 2023-07-07 ; joaquín.sandoval
# modified_on: 2023-07-18 ; guillermo.konig ; Added quarter, start_of_month and end_of_month properties
# modified_on: 2023-07-28 ; guillermo.konig ; Added bussiness day functionality and set_dt argument, to facilitate creating new robot date objects
# modified_on: 2023-08-17 ; vicente.diaz ; added config argument and 

import sys
import os
import re
import logging
import calendar
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
sys.path.append(os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import datetime as dt
from dateutil.relativedelta import relativedelta
from _fmw.fmw_utils import read_config


class RobotDate:
    """
    Class that initiates a date object with several usefull properties

    Arguments:
    - `delta_year`, `delta_month`, `delta_day` (int): modifiers that can be added over the config set ones to edit the datetime from a now date
    - `set_dt` (dt.datetime): If not None the class will ignore offsets and use the given datetime
    - `future` (bool): Determines the time direction of boolean modifier arguments
    - `only_week_day` (bool): Will alter the datetime so that it is not a weekday, moving backwards if `future` is False and fowards if `future` is True
    - `only_bussiness_days` (bool): Will alter the datetime so that it is a bussines day, moving backwards if `future` is False and fowards if `future` is True
    
    For class to work, set up the following config parameters in ENV level:

    "EXECUTION_DAY_OFFSET": 0, // Used to change the date the process will consider as now

    "EXECUTION_YEAR_OFFSET": 0, // Used to change the date the process will consider as now

    "EXECUTION_MONTH_OFFSET": 0,  // Used to change the date the process will consider as now

    Execution offset is meant to be 0 on PRD. 
    """
    def __init__(self, config:dict=None, delta_year:int=0, delta_month:int=0, delta_day:int=0, set_dt:dt.datetime=None, only_week_day:bool=True, only_bussiness_days:bool=False, future:bool=False):
        self.config        = config if config != None else read_config()
        self.config_env    = self.config[self.config["METADATA"]["ENVIRONMENT"].upper()]
        self.config_global = self.config["GLOBAL"]
        # Set datetime by offsets or by a set datetime
        self.default_format = "%Y-%m-%d %H:%M:%S"
        if set_dt is not None:
            self.datetime = set_dt
            self.execution_date = None
        else:
            self.execution_date = self.check_exe_date_in_temp_file()
            self.datetime = (self.execution_date
                                +relativedelta(
                                    years=self.config_env["EXECUTION_YEAR_OFFSET"], 
                                    months=self.config_env["EXECUTION_MONTH_OFFSET"], 
                                    days=self.config_env["EXECUTION_DAY_OFFSET"])
                                +relativedelta(
                                    years=delta_year, 
                                    months=delta_month, 
                                    days=delta_day)
                                )
        # Holiday parameters
        self._holiday_file_path       = os.path.join(self.config["FRAMEWORK"]["PROCESS_DATA"], "robot_date",f"robot_holidays.xlsx")
        self._holiday_file_df         = None 
        self._month_business_days     = None
        self._all_business_days_list  = None
        self._holiday_year_range      = 1
        self._calendar_builder        = CalendarWithHolidays(robot_date=self.datetime, holiday_path=self._holiday_file_path, holiday_year_range= self._holiday_year_range)
        # Change datetime by boolean modifiers
        if only_week_day:
            if self.datetime.weekday() > 4:
                if future:
                    self.datetime = self.datetime + relativedelta(days=7-self.datetime.weekday())
                else:
                    self.datetime = self.datetime + relativedelta(days=4-self.datetime.weekday())
        if only_bussiness_days:
            self.datetime = self.get_previous_or_next_bussiness_day(next=future)

    def check_exe_date_in_temp_file(self) -> dt.datetime:
        """Check if the execution datetime is in a temporal file in process_data"""
        temp_file = os.path.join(self.config["FRAMEWORK"]["PROCESS_DATA"], "_execution_datetime.txt")
        if os.path.exists(temp_file):
            with open(temp_file, "r") as file:
                exe_dti_str = file.read()
                exe_dti     = dt.datetime.strptime(exe_dti_str, self.default_format)
        else:
            with open(temp_file, "w") as file:
                exe_dti     = dt.datetime.now()
                exe_dti_str = exe_dti.strftime(self.default_format)
                file.write(exe_dti_str)
            logging.info(f"Execution date: {exe_dti}")
        return exe_dti
    
    @property
    def set_up(self) -> tuple[str,str,str]:
        """Returns %d, %m , %Y"""
        return self.datetime.strftime("%d"),self.datetime.strftime("%m"),self.datetime.strftime("%Y")
    
    @property
    def set_up_int(self) -> tuple[int,int,int]:
        """Returns day, month, year as a tuple of integer"""
        return self.datetime.day,self.datetime.month,self.datetime.year
    
    # Days area -----------------------------------------------------------------

    @property
    def day(self)->str:
        """01"""
        return self.datetime.strftime("%d")
    
    @property
    def day_short(self)->str:
        """Mon"""
        return self.datetime.strftime("%a")
    
    @property
    def day_short_esp(self)->str:
        """Lun, force allows to "capitalize","upper","lower" the output. Default = None = capitalize... For more info see transform_dates_str()"""
        return self.transform_dates_str(self.datetime.strftime("%A"), to_esp=True)[:3]
    
    def get_day_short_esp(self, force:str=None, lenght:int=3)->str:
        """Lun, force allows to "capitalize","upper","lower" the output. Default = None = capitalize... For more info see transform_dates_str()"""
        esp_day = self.transform_dates_str(self.datetime.strftime("%A"), to_esp=True, force=force)
        lenght = lenght if len(esp_day) > lenght else len(esp_day)
        return esp_day[:lenght]
    
    @property
    def day_full(self)->str:
        """Monday"""
        return self.datetime.strftime("%A")
    
    @property
    def day_full_esp(self)->str:
        """Lunes, force allows to "capitalize","upper","lower" the output. Default = None = capitalize... For more info see transform_dates_str()"""
        return self.transform_dates_str(self.datetime.strftime("%A"), to_esp=True)
    
    def get_day_full_esp(self, force:str=None)->str:
        """Lunes, force allows to "capitalize","upper","lower" the output. Default = None = capitalize... For more info see transform_dates_str()"""
        return self.transform_dates_str(self.datetime.strftime("%A"), to_esp=True, force=force)
    
    @property
    def day_int(self)->int:
        """1"""
        return self.datetime.day
    
    @property
    def total_days_in_month(self) ->int:
        """Returns number of days in the currrent month"""
        total_days = calendar.monthrange(self.year_int, self.month_int)[1]
        return total_days
    
    # Month area -----------------------------------------------------------------

    @property
    def month(self)->str:
        """01"""
        return self.datetime.strftime("%m")
    
    @property
    def month_short(self)->str:
        """Feb"""
        return self.datetime.strftime("%b")
        
    @property
    def month_short_esp(self)->str:
        """Mar"""
        return self.transform_dates_str(self.datetime.strftime("%B"),to_esp=True)[:3]
    
    def get_month_short_esp(self, force:str=None, lenght:int=3)->str:
        """Feb, force allows to "capitalize","upper","lower" the output. Default = None = capitalize... For more info see transform_dates_str()
        lenght allows change the amount of letters. default = 3"""
        esp_month = self.transform_dates_str(self.datetime.strftime("%B"),to_esp=True,force=force)
        lenght = lenght if len(esp_month) > lenght else len(esp_month)
        return esp_month[:lenght]
    
    @property
    def month_full(self)->str:
        """February"""
        return self.datetime.strftime("%B")
    
    @property
    def month_full_esp(self)->str:
        """Febrero, force allows to "capitalize","upper","lower" the output. Default = None = capitalize... For more info see transform_dates_str()"""
        return self.transform_dates_str(self.datetime.strftime("%B"),to_esp=True)
    
    def get_month_full_esp(self, force:str=None)->str:
        """Febrero, force allows to "capitalize","upper","lower" the output. Default = None = capitalize... For more info see transform_dates_str()"""
        return self.transform_dates_str(self.datetime.strftime("%B"),to_esp=True,force=force)

    @property
    def month_int(self)->int:
        """1"""
        return self.datetime.month
    
    # Year area -----------------------------------------------------------------

    @property
    def year(self)->str:
        """0123 as str"""
        return self.datetime.strftime("%Y")
    
    @property
    def year_int(self)->int:
        """123 as int"""
        return self.datetime.year
    
    # Extra -----------------------------------------------------------------
    
    @property
    def delta_days_from_today(self)->int:
        """Amount of Total days from today. Weekends are considered"""
        return (self.datetime - dt.datetime.now()).days

    @property
    def delta_datetime_from_today(self)->dt.datetime:
        """Delta datetime from now"""
        return self.datetime - dt.datetime.now()

    @property
    def to_excel(self)->str:
        """Usefull for adding a datetime using a value directly into win32"""
        return self.datetime.strftime(self.default_format)
    
    @property
    def quarter(self)->int:
        """Returns year quarter the datetime currently is in"""
        if 1<=self.month_int<=3:
            return 1
        elif 4<=self.month_int<=6:
            return 2
        elif 7<=self.month_int<=9:
            return 3
        else:
            return 4
    
    @property
    def start_of_month(self)->dt.datetime:
        """Returns a datetime indicating the first day of the month the process currently is in, resets hours and minutes as well as well"""
        return self.datetime.replace(hour=0, minute=0, second=0, microsecond=0, day=1)

    @property
    def end_of_month(self)->dt.datetime:
        """Returns a datetime indicating the last day of the month the process currently is in, resets hours and minutes as well as well"""
        next_month_first_date = (self.datetime + relativedelta(months=1)).replace(hour=0, minute=0, second=0, microsecond=0, day=1)
        end_of_month = (next_month_first_date + relativedelta(days=-1)).replace(hour=0, minute=0, second=0, microsecond=0)
        return end_of_month

    
    def transform_dates_str(self, string_date:str,to_esp:bool=False,to_eng:bool=False,force:str=None)->str:
        """
        Parameters:
        force (str): Allows to force the output.
        -Valid options are "capitalize","upper","lower".
        -Default is None.
        """
        esp_months_dict = {
            1:'Enero',
            2:'Febrero',
            3:'Marzo',
            4:'Abril',
            5:'Mayo',
            6:'Junio',
            7:'Julio',
            8:'Agosto',
            9:'Septiembre',
            10:'Octubre',
            11:'Noviembre',
            12:'Diciembre'
        }
        esp_days_dict = {
            0:'Lunes',
            1:'Martes',
            2:'Miercoles',
            3:'Jueves',
            4:'Viernes',
            5:'Sabado',
            6:'Domingo'
        }
        eng_months_dict = {
            1:"January",
            2:"February",
            3:"March",
            4:"April",
            5:"May",
            6:"June",
            7:"July",
            8:"August",
            9:"September",
            10:"October",
            11:"November",
            12:"December"
        }
        eng_days_dict = { 
            0:"Monday",
            1:"Tuesday",
            2:"Wednesday",
            3:"Thursday",
            4:"Friday",
            5:"Saturday",
            6:"Sunday"
        }
        def regex_filter(target:str="",target_sub_string:str=""):
            out_str = "("
            for char in target:
                out_str += "["+char.upper()+char.lower()+"]"
            out_str+=")|("
            for char in target_sub_string:
                out_str += "["+char.upper()+char.lower()+"]"
            out_str += ")"
            return out_str
    
        for key_dict in range(0,13):
            if key_dict in esp_days_dict.keys() and key_dict in eng_days_dict.keys():
                if force is None:
                    if to_eng:
                        string_date=string_date.replace(esp_days_dict[key_dict],eng_days_dict[key_dict])
                        string_date=string_date.replace(esp_days_dict[key_dict].lower(),eng_days_dict[key_dict].lower())
                        string_date=string_date.replace(esp_days_dict[key_dict].upper(),eng_days_dict[key_dict].upper())
                    elif to_esp:
                        string_date=string_date.replace(eng_days_dict[key_dict],esp_days_dict[key_dict])
                        string_date=string_date.replace(eng_days_dict[key_dict].lower(),esp_days_dict[key_dict].lower())
                        string_date=string_date.replace(eng_days_dict[key_dict].upper(),esp_days_dict[key_dict].upper())
                elif force == "capitalize":
                    if to_eng:
                        string_date=re.sub(regex_filter(target=esp_days_dict[key_dict],target_sub_string=eng_days_dict[key_dict]),eng_days_dict[key_dict],string_date)
                    elif to_esp:
                        string_date=re.sub(regex_filter(target=esp_days_dict[key_dict],target_sub_string=eng_days_dict[key_dict]),esp_days_dict[key_dict],string_date)
                elif force == "upper":
                    if to_eng:
                        string_date=re.sub(regex_filter(target=esp_days_dict[key_dict],target_sub_string=eng_days_dict[key_dict]),eng_days_dict[key_dict].upper(),string_date)
                    elif to_esp:
                        string_date=re.sub(regex_filter(target=esp_days_dict[key_dict],target_sub_string=eng_days_dict[key_dict]),esp_days_dict[key_dict].upper(),string_date)
                elif force == "lower":
                    if to_eng:
                        string_date=re.sub(regex_filter(target=esp_days_dict[key_dict],target_sub_string=eng_days_dict[key_dict]),eng_days_dict[key_dict].lower(),string_date)
                    elif to_esp:
                        string_date=re.sub(regex_filter(target=esp_days_dict[key_dict],target_sub_string=eng_days_dict[key_dict]),esp_days_dict[key_dict].lower(),string_date)
            if key_dict in esp_months_dict.keys() and key_dict in eng_months_dict.keys():
                if force is None:
                    if to_eng:
                        string_date=string_date.replace(esp_months_dict[key_dict],eng_months_dict[key_dict])
                        string_date=string_date.replace(esp_months_dict[key_dict].lower(),eng_months_dict[key_dict].lower())
                        string_date=string_date.replace(esp_months_dict[key_dict].upper(),eng_months_dict[key_dict].upper())
                    elif to_esp:
                        string_date=string_date.replace(eng_months_dict[key_dict],esp_months_dict[key_dict])
                        string_date=string_date.replace(eng_months_dict[key_dict].lower(),esp_months_dict[key_dict].lower())
                        string_date=string_date.replace(eng_months_dict[key_dict].upper(),esp_months_dict[key_dict].upper())
                elif force == "capitalize":
                    if to_eng:
                        string_date=re.sub(regex_filter(target=esp_months_dict[key_dict],target_sub_string=eng_months_dict[key_dict]),eng_months_dict[key_dict],string_date)
                    elif to_esp:
                        string_date=re.sub(regex_filter(target=esp_months_dict[key_dict],target_sub_string=eng_months_dict[key_dict]),esp_months_dict[key_dict],string_date)
                elif force == "upper":
                    if to_eng:
                        string_date=re.sub(regex_filter(target=esp_months_dict[key_dict],target_sub_string=eng_months_dict[key_dict]),eng_months_dict[key_dict].upper(),string_date)
                    elif to_esp:
                        string_date=re.sub(regex_filter(target=esp_months_dict[key_dict],target_sub_string=eng_months_dict[key_dict]),esp_months_dict[key_dict].upper(),string_date)
                elif force == "lower":
                    if to_eng:
                        string_date=re.sub(regex_filter(target=esp_months_dict[key_dict],target_sub_string=eng_months_dict[key_dict]),eng_months_dict[key_dict].lower(),string_date)
                    elif to_esp:
                        string_date=re.sub(regex_filter(target=esp_months_dict[key_dict],target_sub_string=eng_months_dict[key_dict]),esp_months_dict[key_dict].lower(),string_date)        
        return string_date

    def short_esp_date_to_num(self, str_date:str, point_abreviation:bool=True):
        esp_months_dict = {
            "01":'[Ee][Nn][Ee]\.?',
            "02":'[Ff][Ee][Bb]\.?',
            "03":'[Mm][Aa][Rr]\.?',
            "04":'[Aa][Bb][Rr]\.?',
            "05":'[Mm][Aa][Yy]\.?',
            "06":'[Jj][Uu][Nn]\.?',
            "07":'[Jj][Uu][Ll]\.?',
            "08":'[Aa][Gg][Oo]\.?',
            "09":'[Ss][Ee][Pp][Tt]\.?|[Ss][Ee][Pp]\.?',
            "10":'[Oo][Cc][Tt]\.?',
            "11":'[Nn][Oo][Vv]\.?',
            "12":'[Dd][Ii][Cc]\.?'
        }
        esp_months_dict_no_point = {
            "01":'[Ee][Nn][Ee]',
            "02":'[Ff][Ee][Bb]',
            "03":'[Mm][Aa][Rr]',
            "04":'[Aa][Bb][Rr]',
            "05":'[Mm][Aa][Yy]',
            "06":'[Jj][Uu][Nn]',
            "07":'[Jj][Uu][Ll]',
            "08":'[Aa][Gg][Oo]',
            "09":'[Ss][Ee][Pp][Tt]|[Ss][Ee][Pp]',
            "10":'[Oo][Cc][Tt]',
            "11":'[Nn][Oo][Vv]',
            "12":'[Dd][Ii][Cc]'
        }
        if point_abreviation:
            for key_dict in esp_months_dict.keys():
                str_date=re.sub(esp_months_dict[key_dict],key_dict,str_date)
        else:
            for key_dict in esp_months_dict_no_point.keys():
                str_date=re.sub(esp_months_dict_no_point[key_dict],key_dict,str_date)
        return str_date
    
    @property
    def holidays_df(self) -> pd.DataFrame:
        """Returns holiday dataframe if posible"""
        try: 
            if self._holiday_file_df is None:
                self._calendar_builder.build_calendar()
                self._holiday_file_df = pd.read_excel(self._holiday_file_path, sheet_name="base")
            return self._holiday_file_df
        except:
            logging.info("Failed to create holiday dataframe returning empty one")
            return pd.DataFrame(columns=[["HOLIDAY","DATE"]])
        
    @property
    def holiday_list(self) ->list[pd.DataFrame]:
        """Returns a list of all holidays"""
        return self.holidays_df["DATE"].tolist()

    @property
    def is_holiday(self)->bool:
        """Returns a boolean indicating if robot date is holiday"""
        df_holidays = self.holidays_df[self.holidays_df['HOLIDAY'] == True]
        df_verify   = df_holidays[df_holidays['DATE'].isin([self.datetime.strftime("%Y/%m/%d")])]
        if len(df_verify) > 0:
            holiday =  True
        else:
            holiday = False
        return holiday
    
    def is_date_holiday(self, date:dt.datetime=None)->bool:
        """Returns a boolean indicating if date is holiday"""
        if date is None:
            date = self.datetime
        df_holidays = self.holidays_df[self.holidays_df['HOLIDAY'] == True]
        df_verify   = df_holidays[df_holidays['DATE'].isin([date.strftime("%Y/%m/%d")])]
        if len(df_verify) > 0:
            holiday =  True
        else:
            holiday = False
        return holiday
    
    @property
    def is_bussines_day(self) -> bool:
        """Returns a boolean indicating if robot date is a business day"""
        week_day     = self.datetime.weekday() # Day of week: Monday=0... Sunday=6
        holiday      = self.is_holiday # Verify is not a holiday
        # If the date is not Saturday and Sunday and Holiday, then is Bussines day
        if week_day!=5 and week_day!=6 and holiday==False:
            bussines_day = True
        else:
            bussines_day = False
        return bussines_day

    def is_date_bussines_day(self, date:dt.datetime=None) -> bool:
        """Returns a boolean indicating if date is a business day"""
        if date is None:
            date = self.datetime
        week_day     = date.weekday() # Day of week: Monday=0... Sunday=6
        holiday      = self.is_date_holiday(date=date) # Verify is not a holiday
        # If the date is not Saturday and Sunday and Holiday, then is Bussines day
        if week_day!=5 and week_day!=6 and holiday==False:
            bussines_day = True
        else:
            bussines_day = False
        return bussines_day
    
    @property
    def month_business_days(self) ->list[dt.datetime]:
        """ List of all datetime Bussines days in month. """
        if not self._month_business_days:
            self._month_business_days = list()
            for i in range(1, self.total_days_in_month + 1):
                date         = dt.datetime(year=self.year_int, month=self.month_int, day=i)
                bussines_day = self.is_date_bussines_day(date=date)
                if bussines_day == True:
                    self._month_business_days.append(date)
        return self._month_business_days
    
    @property
    def all_business_days(self) ->list[dt.datetime]:
        """Returns a list of all the business days in the previous, current and next year"""
        if not self._all_business_days_list:
            self._all_business_days_list = list()
            for loop_year in range(self.year_int-self._holiday_year_range, self.year_int+ self._holiday_year_range+1):
                for month in range(1,13):
                    for day in range(1, calendar.monthrange(loop_year, month)[1]):
                        date = dt.datetime(day= day, month =month, year=loop_year)
                        bussines_day = self.is_date_bussines_day(date=date)
                        if bussines_day == True:
                            self._all_business_days_list.append(date)
        return self._all_business_days_list
        
    @property
    def end_of_bussiness_month(self)->dt.datetime:
        """Returns the last bussiness day of the current month"""
        return self.month_business_days[-1]

    @property
    def start_of_bussiness_month(self)->dt.datetime:
        """Returns the first bussiness day of the current month"""
        return self.month_business_days[0]
    
    def get_previous_or_next_bussiness_day(self, next:bool=False):
        """If next is set to True it will get the next bussiness day instead of the last"""
        if next:
            modifier = -1
        else:
            modifier = 1
        # Checks up to a year back for a business day
        for idx in range(0,366):
            date = self.datetime+relativedelta(days=modifier*idx)
            business_day = self.is_date_bussines_day(date=date)
            if business_day:
                return date

    def is_date_in_first_or_last_month_bussiness_days(self, day_flag:int, first:bool=False):
        """
        Returns boolean indicating if robot date is within a range of days set by `day_flag` and if its the first or last days set by `first`

        Examples: 

        If `day_flag` is set to 5 and `first` is set to True, if the robot date is within the first 5 business days it will return True

        If `day_flag` is set to 5 and `first` is set to False, if the robot date is within the last 5 business days it will return True
        """
        if first:
            range_bussiness_days = self.month_business_days[0:day_flag+1]
            if any([day for day in range_bussiness_days if self.datetime.strftime("%Y%m%d") == day.strftime("%Y%m%d")]):
                return True
            return False
        else:
            range_bussiness_days = self.month_business_days[-day_flag:]
            if any([day for day in range_bussiness_days if self.datetime.strftime("%Y%m%d") == day.strftime("%Y%m%d")]):
                return True
            return False
        
    def __repr__(self):
        return self.to_excel


class CalendarWithHolidays():
    def __init__(self, robot_date:dt.datetime, holiday_path:str, holiday_year_range:int=1):
        self.year             = robot_date.year
        self.day              = int(robot_date.day)
        self.list_holidays    = list()
        self.holiday_year_range = holiday_year_range
        # These variables already exist by inheritance
        self.url_holidays      = 'https://www.officeholidays.com/countries/chile/{0}'
        self.download_path     = os.path.dirname(holiday_path)
        self.file_path         = holiday_path
        self.keep_web_alive    = False
        self.excecute_headless = True # True = Execute in background
        

    def init_edge_driver(self):        
        edge_options = Options()
        os.makedirs(self.download_path, exist_ok=True)
        prefs = {"profile.default_content_settings.popups": 0,    
                 "download.default_directory": self.download_path}
        if self.excecute_headless:
            edge_options.add_argument('--headless=new')
        edge_options.add_argument("--start-maximized")
        edge_options.add_argument('--no-sandbox')
        edge_options.add_argument('--disable-dev-shm-usage')
        edge_options.add_argument('--ignore-certificate-errors')
        edge_options.add_argument('--ignore-ssl-errors')
        edge_options.add_argument("--disable-popup-blocking")
        edge_options.add_experimental_option('excludeSwitches', ['enable-logging', 'disable-popup-blocking'])
        edge_options.add_experimental_option("detach", self.keep_web_alive)
        edge_options.add_experimental_option("prefs", prefs)
        # Open website
        self.driver = webdriver.Edge(options=edge_options, keep_alive=self.keep_web_alive)
   
    def open_website(self, url:str):
        """ Open website. """
        logging.info(f"Open the website {url}...")
        self.driver.get(url)

    def get_holidays_dates(self, year:str):
        """ Get the holidays of chile of the year. """
        logging.info(f"Get holidays dates from Chile year {year}...")
        selector_date = '//*[@id="wrapper"]/div[5]/div[2]/table[1]/tbody/tr/td[2]'
        list_dates = self.driver.find_elements(by=By.XPATH, value=selector_date)
        for i in range(len(list_dates)):
            date_i = list_dates[i].text
            datetime_i = dt.datetime.strptime(f"{year} {date_i}", "%Y %b %d")
            self.list_holidays.append({"DATE": datetime_i.strftime("%Y/%m/%d"), "HOLIDAY": True})

    def save_to_excel(self):
        """ Save holidays in Excel. """
        df_holidays = pd.DataFrame(self.list_holidays)
        df_holidays.to_excel(self.file_path, "base")
        logging.info(f"Excel saved in path: {self.file_path}")

    def close_website(self):
        """ Close website with Selenium. """
        logging.info(f"Closing website {self.url_holidays}")
        self.driver.quit()

    def verify_existence(self):
        """ If the file Holydays does not exist or the current date is within the first 
        week of the month or the file has no data for the desiered year, the file will be download from web. """
        if os.path.exists(self.file_path):
            df_holidays = pd.read_excel(self.file_path)
            if not df_holidays.empty:
                dates = df_holidays["DATE"].to_list()
                if not any([date for date in dates if dt.datetime.strptime(date, "%Y/%m/%d").year == self.year]):
                    return False
                # if self.day in range(1,8):
                #     return False
                else:
                    return True
            else:
                return False
        return False

    def build_calendar(self):
        try: 
            if not self.verify_existence():
                logging.info("Getting office holidays")
                self.init_edge_driver()
                # Get all years within the range given
                for year_mod in range(-self.holiday_year_range, self.holiday_year_range+1):
                    year = self.year+year_mod
                    self.open_website(url=self.url_holidays.format(year))
                    self.get_holidays_dates(year=year)
                self.save_to_excel()
                self.close_website()
                logging.info("Finished getting office holidays")
            else:
                pass
        except Exception as error:
            logging.warning(f"Failed to get office holidays due to error: {error}")


if __name__ == "__main__":
    exe_date = RobotDate()
    print(exe_date)