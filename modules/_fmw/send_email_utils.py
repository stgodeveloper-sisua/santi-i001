# file_name:   send_email_utils.py
# created_on:  2023-03-29 ; guillermo.konig
# modified_on: 2023-08-04 ; guillermo.konig ; Added reply all, the option to send as another user and image_to_html

import os
import logging
import warnings
import base64
import win32com.client as win32com
import pandas as pd


def read_recipients_file(recipients_file_path: str
                         , mail_type: str
                         , environment: str="dev"
                         , sheet: str= "base"):
    """
    Reads recpients file and returns three lists to, cc, and bcc.
    Lists contain the emails that are assigned at the desiered environment and mail_type.
    For an email to be considered it must be TRUE in the email type and of the desiered environment.
    The email will be sent as to, cc or bcc depending on the value of recipient_type column.
    INPUTS
    -recipients_file_path (str): File path to the recipients file, must be of type  (found in template found with this utils)
    -mail_type (str): Column that indicates the type or email, must be a column in recipients file (The rows marked as TRUE mean the user should be sent if it matches other conditions.)
    -environment (str): Must be either 'dev' or 'prod' indicates the environment to filter accounts (Default is 'dev')
    -sheet(str): Sheetname of recipients file (Default is 'base')

    Created By: Guillermo König
    Created At:  02/03/2023
    Modified At: 02/03/2023
    """
    if not os.path.exists(recipients_file_path):
        raise Exception("Recipients file does not exist")
    warnings.simplefilter(action='ignore', category=UserWarning)
    mail_recipients = pd.read_excel(recipients_file_path, sheet)
    warnings.resetwarnings()
    mail_columns    = mail_recipients.columns
    if not mail_type in mail_columns:
        raise Exception("Mail type indicated could not be found in recipients file")
    # Dataframe masks
    desiered_env = mail_recipients["environment_level"].str.upper() == environment.upper() 
    true_to_send = mail_recipients[mail_type] == "TRUE"
    bool_to_send = mail_recipients[mail_type] == True
    verdadero_to_send = mail_recipients[mail_type] == "VERDADERO"
    # Assign recipient lists
    to_recipient_type  = mail_recipients["recipient_type"]=="to"
    cc_recipient_type  = mail_recipients["recipient_type"]=="cc"
    bcc_recipient_type = mail_recipients["recipient_type"]=="bcc"
    To = list(mail_recipients[desiered_env & (bool_to_send | verdadero_to_send | true_to_send) & to_recipient_type]["email_address"])
    Cc = list(mail_recipients[desiered_env & (bool_to_send | verdadero_to_send | true_to_send) & cc_recipient_type]["email_address"])
    Bcc = list(mail_recipients[desiered_env & (bool_to_send | verdadero_to_send | true_to_send) & bcc_recipient_type]["email_address"])
    return To, Cc, Bcc

def image_to_html(image_path:str, height:int=None, width:int=None):
    try:
        with open(image_path, "rb") as image_file:
            image_data  = image_file.read()
            base64_data = base64.b64encode(image_data).decode("utf-8")
            mime_type = "image/"+ image_path.split(".")[-1].lower()
            data_url  = f"data:{mime_type};base64,{base64_data}"
            if height:
                height_html = f' height="{height}"'
            else:
                height_html = ''
            if width:
                width_html = f' width="{width}"'
            else:
                width_html = ''
            html  = f'<img src="{data_url}" alt="robot_body_img">{height_html}{width_html}>'
            return html
    except IOError as e:
        logging.warning(f"Error reading the image file {e}")
        return "" 

def df2html(df: pd.DataFrame
            , width=80
            , font="open sans"
            , border_color="#000"
            , header_bg_color="#012351"
            , header_font_color="#ffffff"
            , first_col_alignment="left"
            , default_alignment="left"
            , even_rows_color="#D5D5D5"
            , odd_rows_color="#ffffff"
            , font_size=14):
    """
    Transforms a dataframe to an html that can be used to show the table in an email
    Created By: Vicente Diaz
    Created At:  
    Modified At: 
    """
    # Table initial format
    table_format  = f"<table cellpadding='5' style='border: 2px solid {border_color};" \
                    f"border-collapse: collapse; font-family:{font};font-size:{font_size}px;"\
                    f"margin-left:auto;margin-right:auto;' width='{width}%'>\n"
    # Add table headers
    header_format = f"<thead> <tr style='border: 2px solid {border_color} ;background-color: " \
                    f"{header_bg_color}; color:{header_font_color};'>\n"
    header_data   = header_format
    for idx, header in enumerate(df.columns):
        if idx == 0:
            text_align = first_col_alignment
        else:
            text_align = default_alignment
        header_data += f"<th style='text-align: {text_align};'> {header} </th>\n"
    header_data += "</tr></thead>\n"
    # Add table body
    body_data = "<tbody>\n"
    for idx_row, row in df.iterrows():
        if (idx_row+1) % 2 == 1:
            row_data = f"<tr style='background-color: {odd_rows_color};'> "
        else:
            row_data = f"<tr style='background-color: {even_rows_color};'> "
        for idx_value, value in enumerate(row.values):
            if idx_value == 0:
                text_align = first_col_alignment
            else:
                text_align = default_alignment
            row_data += f"<td style='text-align: {text_align};'> {value} </td>\n"
        row_data += "</tr>\n"
        body_data += row_data
    body_data += "</tbody>\n"
    # Join everything
    html_df = table_format + header_data + body_data + "</table>\n"
    return html_df
            
def get_account_obj_by_email(outlook_session, account_email):
    accounts = outlook_session.Session.Accounts
    for account in accounts:
        if account.DisplayName == account_email or account.SmtpAddress == account_email:
            return account
    return None

def send_email(subject:str
               , to:list
               , body_file:str
               , wrapper_file:str
               , environment:str = "dev"
               , body_fields: list = []
               , wrapper_fields: list = []
               , cc:  list   = []
               , bcc: list   = []
               , attachments = []
               , sender_account:str = None):
    """
    Sends email with logged account as defined by parameters
    -subject (String)              : Email subject
    -to (List[String])             : List of to level recipients
    -cc (List[String])             : List of cc level recipients
    -bcc (List[String])            : List of bcc level recipients
    -body_file (String)            : Filepath to text file containning body
    -body_fileds (List[String])    : List of values that will be formatted into body_file text
    -wrapper_file (String)         : Filepath to email wrapper html
    -wrapper_fields (List[String]) : List of values that will be formatted into wrapper html
    -environment (String)          : Used to determine if the email was sent as dev or qa, if so adds [TEST] as prefix to subject    
    
    Created By: Guillermo König
    Created At:  02/03/2023
    Modified At: 02/03/2023
    """
    logging.info(f"Sending email: Subject: {subject}")
    outlook = win32com.Dispatch('outlook.application') 
    mail = outlook.CreateItem(0)
    if environment.upper()=="DEV" or environment.upper()=="QAS":
        subject = "[TEST] "+subject
    mail.Subject = subject
    # Choose extra account if given
    if sender_account:
        account = get_account_obj_by_email(outlook_session=outlook, account_email=sender_account)
        if account:
            logging.info(f"Sending email by: {account.DisplayName}")
            mail.SentOnBehalfOfName = sender_account
            #mail._oleobj_.Invoke(*(64209, 0, 8,0,account))
        else:
            raise Exception(f"Failed to find account {sender_account} in outlook session")
    # Add recipients as strings
    logging.info(f"Recipients -> to: {to} | cc: {cc} | bcc: {bcc}")
    if len(to+cc+bcc)<0:
        logging.error("No recipients given for email, unable to send")
        return False
    if len(to)>0:
        mail.To = ";".join(to)
    if len(cc)>0:
        mail.Cc = ";".join(cc)
    if len(bcc)>0:
        mail.Bcc = ";".join(bcc)
    # Read and create body message
    with open(body_file, "r", encoding="UTF-8") as b_file:
        body_msg = b_file.read().format(*body_fields)
    wrapper_fields = [body_msg] + wrapper_fields
    # Read wrapper template
    with open(wrapper_file, "r", encoding="UTF-8") as w_file:
        wrapper = w_file.read().format(*wrapper_fields)
    email_msg     = wrapper
    mail.HTMLBody = email_msg
    # Attachments
    logging.info(f"Attachments: {len(attachments)}")
    for attachment in attachments:
        if os.path.exists(attachment):
            mail.Attachments.Add(attachment)
        else:
            logging.warning(f"Attachment {attachment} could not be found")
    mail.Send()
    logging.info("Email sent successfuly")
    return True

def reply_all_email(mail
            , body_file:str
            , wrapper_file:str
            , body_fields: list = []
            , wrapper_fields: list = []
            , to:  list   = []
            , cc:  list   = []
            , bcc: list   = []
            , attachments = []
            , sender_account:str = None):

    def get_message_recipients(message):
        def remove_dupl_str(strs_joined:str, joiner:str):
            lst_strs = strs_joined.split(joiner)
            remv_dupli = list(dict.fromkeys(lst_strs))
            strs_out = joiner.join(remv_dupli)
            return strs_out
        recipients = message.Recipients
        to = ""
        cc = ""
        bcc = ""
        for recipient in recipients:
            recipient_type = recipient.Type
            display_type = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            if str(recipient_type) == "1":
                to += f";{display_type}"
            if str(recipient_type) == "2":
                cc += f";{display_type}"
            if str(recipient_type) == "3":
                bcc += f";{display_type}"
        return remove_dupl_str(to,";"), remove_dupl_str(cc,";"), remove_dupl_str(bcc,";")

    outlook = win32com.Dispatch('outlook.application') 
    mail = mail.ReplyAll()
    logging.info(f"Sending email: Subject: {mail.Subject}")
    # Choose extra account if given
    if sender_account:
        account = get_account_obj_by_email(outlook_session=outlook, account_email=sender_account)
        if account:
            logging.info(f"Sending email by: {account.DisplayName}")
            mail.SentOnBehalfOfName = sender_account
            #mail._oleobj_.Invoke(*(64209, 0, 8,0,account))
        else:
            raise Exception(f"Failed to find account {sender_account} in outlook session")
    # Add recipients as strings
    logging.info(f"Extra Recipients -> to: {to} | cc: {cc} | bcc: {bcc}")   
    old_mail_to, old_mail_cc, old_mail_bcc = get_message_recipients(mail)
    if len(to)>0:
        mail.To = ";".join(to)
        # Add all To to cc so the mail state can be tracked
        cc = (f"{';'.join(cc)};{old_mail_to}").split(";")
    if len(cc)>0:
        mail.Cc = ";".join(old_mail_cc.split(";")+cc) 
    if len(bcc)>0:
        mail.Bcc = ";".join(old_mail_bcc.split(";")+bcc)
    # Read and create body message
    with open(body_file, "r", encoding="UTF-8") as b_file:
        body_msg = b_file.read().format(*body_fields)
    wrapper_fields = [body_msg] + wrapper_fields
    # Read wrapper template
    with open(wrapper_file, "r", encoding="UTF-8") as w_file:
        wrapper = w_file.read().format(*wrapper_fields)
    email_msg     = wrapper
    mail.HTMLBody = email_msg + mail.HTMLBody
    # Attachments
    logging.info(f"Attachments: {len(attachments)}")
    for attachment in attachments:
        if os.path.exists(attachment):
            mail.Attachments.Add(attachment)
        else:
            logging.warning(f"Attachment {attachment} could not be found")
    mail.Send()
    logging.info("Email sent successfuly")
    return True