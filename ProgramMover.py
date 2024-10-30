# Project Name: Program Mover
# Author: Vinh Huynh
# Date: 07/17/2024
# Version: 0000

import os
import sys
import datetime
import time
# Handles JSON
import json

# Used to access excel file
import pandas as pd
import win32com.client
# Handles the file copying
import shutil 

# Handles the emailing and account information
import smtplib
from email.message import EmailMessage
# Accesses file for account information
from decouple import config

#----------------------------------------For Testing----------------------------------------#
# Switch Variable Below to True for testing/ False for production
testing = True
#-------------------------------------------------------------------------------------------#



#-------------------------------------Email Information-------------------------------------#

"""Change Email Information Here"""
if testing == True:
    recipient =  ""
elif testing == False:
    recipient = [""]

cc_recipient = ""
error_recipient = [""]

#-----------------------------------Change Time/Date Here-----------------------------------#

"""Check File Modification Date"""

#check_date = datetime.datetime(2023,1,1)
check_src_date = datetime.datetime.now() - datetime.timedelta(days = 1095) # (3years) Update date here for checking if source program is older than certain timeframe
"""Check Time"""
check_time = 360 # (hrs) Update to change how often the program runs check on same item+machine combo

#----------------------------------Item/Program Information---------------------------------#

"""Change Item Information Here"""
bur_names_list = ["DOUBLECUT", "SINGLECUT","ALUMACUT","DIEMILL","TIRE BUR","FIBERGLASS ROUTER", "100-", "IND", "DEN"]
bur_product_code_list = ["BR","DB","DFP","RZ"]
ignore_product_code_list = ["US","ST","RS","CR","COAT"]
checking_item_list = []

#------------------------------------Machine Information------------------------------------#
"""Change Machine Information Here"""
if testing == True: 
    rollomatic_src_path = "\\Temp-folder"
    rollomatic_dst_path = "\\Test1"
    anca_src_path = "\\Temp-Folder Anca"
    anca_dst_path = "\\Anca"
    print("🟡 IN TEST MODE 🟡")

elif testing == False:
    rollomatic_src_path = "\\Temp-folder"
    rollomatic_dst_path = "\\machines"

    anca_src_path = "\\Temp-Folder Anca"
    anca_dst_path = "\\Anca"

# Variables to stored allowed and not allowed machines ex "R" = Rollomatic machines
# Todo Add prep if needed
allowed_machines_dict = {"R": "Rollomatic","A": "Anca", "W": "Walter"}

#------------------------------------Logging Information------------------------------------#

# Stores information and error messages about each execution
program_log_path = "\\Program_Log.txt"
error_log_path = "\\Error_Program_Log.txt"

log_result = ""
error_result = ""
smtp_error = ""
copy_status = ""
email_sent = ""
#--------------------------------------Other File Paths--------------------------------------#
jobs_excel_file = r"\Jobs_Created.xlsx"
machine_excel_file = "\\Machine_Location.xlsx"
if testing == True:
    data_json_path = "t\\Checked Item.json"
elif testing == False:
    data_json_path = "\\Checked Item.json"
    
scheduler_excel_path = "\\Current_Scheduler.xlsx"
machine_file_path_list = []

# Main function
def main():

    global item_program, file_src, item_description, product_code, assigned_machine, checking_item_dict, item_key, job_date, copied_program_list, bur_check, copy_status
    item_program = ""
    item_description = ""
    product_code = ""
    job_date = ""
    assigned_machine = ""
    date_format = "%m-%d-%Y %I:%M:%S"
    # Unformated to use in datetime calculation
    current_date = datetime.datetime.now()
    copied_program_list = []

    #Variables to store email lists
    global walter_email_list, no_vgp_email_list, no_program_email_list
    walter_email_list = []
    no_vgp_email_list = []
    no_program_email_list = []
    # Function to read excel for latest jobs and retrieving item information
    # Used with query from infor/sql in excel
    # For future use when assigned machine is available in sql
    def read_jobs_excel():
        if not os.path.exists(jobs_excel_file):
            print("File Not Found")
            return "Error"
        else:
            job_item = pd.read_excel(jobs_excel_file)
            global checking_item_list, empty_list
            empty_list = [] # To-Do Account for empty rows
            checking_item_list = []
            current_date = datetime.date.today().strftime("%m-%d-%Y")
            for index, row in job_item.iterrows():
                
                if pd.isnull(row["item"]) or pd.isnull(row["job_date"]) or pd.isnull(row["product_code"]):
                    empty_list.append(row["job"])
                    continue
                else:
                    job_date_format = pd.to_datetime(row["job_date"]).strftime("%m-%d-%Y")
                if row["product_code"] in ignore_product_code_list:
                    continue
                if job_date_format < current_date:
                    continue 
                else:
                    checking_item_list.append((row["item"], row["job_date"], row["product_code"], row["uf_coitemnotes"]))
                    # To-Do for future use when assigned machine is avaialable , row["assigned_machine"])
    
    # Function to check assigned machine with current old scheduler
    def read_old_scheduler():
        global checking_item_list, job_checking
        
        if not os.path.exists(scheduler_excel_path):
            print("File Not Found")
            return "Error"
        else:
            
    # Refresh Excel queries
            excel = win32com.client.Dispatch("Excel.Application")
            print("Refreshing Excel Connection ↺")
            wb = excel.Workbooks.Open(scheduler_excel_path)
            excel.DisplayAlerts = False
            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()
            wb.Save()
            wb.Close()
            del(wb)
            excel.Quit()
            print("Excel Refreshed ✓\n")

            df = pd.read_excel(scheduler_excel_path, sheet_name="Append1")
            job_checking = 0
            machine_list = ["R","A","W","P"]
            def check_job(job):
                global job_assinged_machine, machine_status, job_checking
                # Assign machine base on order pulled from excel for now
                if job.startswith(tuple(machine_list)):
                    if "-" not in job:
                        job = job[0] + "-" + job[1:3]
                    if len(job) > 4:
                        job = job[0:4].upper()
                    # Will Assigned machine based on order pulled from excel for now
                    job_assinged_machine = job
                    machine_status = machine_status_row
                if job.startswith("J0"):
                    job_checking = job_checking + 1
                    checking_item_list.append((item_row, description_row, job_assinged_machine))

            for index, row in df.iterrows():
                job = str(row["Job1"]).upper()
                item_row = str(row["Item1"]).upper()
                description_row = str(row["Description1"]).upper()
                machine_status_row = str(row["Status1"]).upper()
                complete_row = str(row["Complete1"])
                if pd.isnull(job):
                    continue
                if job == "JOB":
                    continue
                if complete_row == "1" or item_row == "JOB NOT FOUND" or machine_status_row == "DOWN":
                    continue
                else:
                    check_job(job)

            for index, row in df.iterrows():
                job = str(row["Job2"]).upper()
                item_row = str(row["Item2"]).upper()
                description_row = str(row["Description2"]).upper()
                machine_status_row = str(row["Status2"]).upper()
                complete_row = str(row["Complete2"])
                if pd.isnull(job):
                    continue
                if job == "JOB":
                    continue
                if complete_row == "1" or item_row == "JOB NOT FOUND" or machine_status_row == "DOWN":
                    continue
                else:
                    check_job(job)
    # Function to check if item is a bur
    def check_bur():
        global item_type, accepted_files_list, product_code # To-Do Remove product_code future when not using old scheduler
        product_code = ""
        item_type = ""
        bur_name_ending = ""
        accepted_files_list = []
        # Checks if item is a bur has to match only one criteria
        if item_program.startswith("S"):#ToDO Remove in future when not using old scheduler
            if len(item_program.split("-",1)[0]) in [3,2] and "DM" in item_program.upper():
                product_code = "BR" # To-Do Remove product_code future when not using old scheduler
        if item_program.startswith(tuple(bur_names_list)) or product_code in bur_product_code_list or any(name.lower() in item_description.lower() for name in bur_names_list):
            item_type = "Bur"
            accepted_files_list = [".vgp"]
            if product_code != "SP" and not item_program.startswith(tuple(bur_names_list)):
                # Split out coating
                if "-" in item_program:
                    bur_name_split = item_program.split("-", 3)
                    bur_name_count = item_program.count("-")
                    bur_type = bur_name_split[0].upper()
                    bur_cut =  bur_name_split[1].upper()
                    if bur_name_count >= 2:
                        if not bur_name_split[2].isdigit() and len(bur_name_split[2]) > 1:
                            bur_name_ending = bur_name_split[2]
                            if "-" in bur_name_ending:
                                bur_name_ending = bur_name_ending.split("-", 1)[0]
                else:
                    joined_bur = item_program
                    return joined_bur
                if "R" in bur_cut.upper():
                    index = bur_cut.index("R")
                    bur_cut = bur_cut[:index] + bur_cut[index + 1:]
                if "X" in bur_cut.upper():
                    index = bur_cut.index("X")
                    bur_cut = bur_cut[:index] + bur_cut[index + 1:]
                # Checks for burs with "L6, L120, etc."
                if "L" in bur_cut.upper():
                    index = bur_cut.index("L")
                    bur_cut = bur_cut[:index] + bur_cut[index + 1:]
                    while index < len(bur_cut) and bur_cut[index].isdigit():
                        bur_cut = bur_cut[:index] + bur_cut[index + 1:]
                
                joined_bur = bur_type + "-" + bur_cut
                if bur_name_ending != "":
                    joined_bur = joined_bur + "-" + bur_name_ending
                return joined_bur
            else:
                item_type = "Special Bur"
        else:   
            accepted_files_list = [".vgp"]
            item_type = "Not Bur"
        return item_type
    # function to add to JSON
    def read_json(checking_item):
        global email_sent, check_again
        check_again = False
        email_sent = False
        if os.path.exists(data_json_path):
            # Open JSON file to read/write
            with open(data_json_path,'r+') as file:
                # Function to update JSON
                def json_file_access(file, file_data):
                    file.seek(0)
                    json.dump(file_data, file, indent = 4)
                    file.truncate()
                    file.close()
                # Handle if file is empty
                if os.path.getsize(data_json_path) == 0:
                    json_default = {"checked_item_details":[]}
                    # Converts python object to JSON string
                    json.dump(json_default, file, indent = 4)
                if os.path.getsize(data_json_path) > 0:
                    file_data = json.load(file)
                    # Accessing dictionary values stored within list
                    for key_dict in file_data["checked_item_details"]: # file_data type = list
                        if key_dict.get("key") == item_key: # Key_dict type = dict
                            if copy_status != "":
                                key_dict.update({"Copy Status": copy_status})
                                key_dict.update({"Date": current_date.strftime(date_format)})
                                json_file_access(file, file_data)
                                return
                            # Prevents checking program if it has already been checked within specified time
                            data_date = datetime.datetime.strptime(key_dict.get("Date"), date_format)
                            if current_date - data_date > datetime.timedelta(hours = check_time):
                                # Prevents resending of email for checked items that have not been completed
                                key_dict.update({"Date": current_date.strftime(date_format)})
                                check_again = True
                                json_file_access(file, file_data)
                                return "Continue Check"
                            elif current_date - data_date < datetime.timedelta(hours = check_time):
                                if key_dict.get("Copy Status") == "Email Sent":
                                    if key_dict.get("Assigned Machine").startswith("W"):
                                        return "Already Checked"
                                    email_sent = True
                                    return "Continue Check"
                                if key_dict.get("Copy Status") == "":
                                    return "Continue Check"
                                return "Already Checked"
                    if not any(key_dict.get("key") == item_key for key_dict in file_data["checked_item_details"]):
                        file_data["checked_item_details"].append(checking_item)
                        json_file_access(file, file_data)
                        return "Continue Check"    
        elif not os.path.exists(data_json_path):
            error_log(f"{data_json_path}\n--File Does Not Exist")
            return "Error" 
    # Use for future when pulling info from INFOR/SQL
    """
    # Call to read_jobs_excel function using info pulled from infor/Sql
    if read_jobs_excel() == "Error":
        program_status = f"Program Status: Error Reading Jobs Excel"
        recipient = "caduser1@mastercuttool.com"
        send_email("Error Reading Jobs Excel", recipient, cc_recipient, program_status)
    if checking_item_list == []:
        program_status = f"Program Status: No Items Read From Jobs Excel"
        recipient = "caduser1@mastercuttool.com"
        send_email("No items to check", recipient, cc_recipient, program_status)
    else:
        for item_info in checking_item_list:
            item_program, job_date, product_code, item_description = item_info"""
    global start_time
    start_time=time.time()
    # Call to read_old_scheduler function
    if read_old_scheduler() == "Error":
        program_status = f"Program Status: Error Reading Old Scheduler Excel"
        send_email("Error Reading Old Scheduler Excel Excel", "caduser1@mastercuttool.com", cc_recipient, program_status)
    if checking_item_list == []:
        program_status = f"Program Status: No Items Read From Old Scheduler Excel"
        send_email("No items to check", "caduser1@mastercuttool.com", cc_recipient, program_status)
    else:
        count = 0
        for item_info in checking_item_list:
            item_program, item_description, assigned_machine = item_info
            if "P" in assigned_machine:
                continue
            # Python object to be appended to JSON
            # To-Do Will need to add more information on the info in dict if needed like start and end time 
            checking_item_dict = {"key": f"{item_program}{assigned_machine}",
                                "Checked Item": f"{item_program}",
                                "Assigned Machine": f"{assigned_machine}",
                                "Date": f"{current_date.strftime(date_format)}",
                                "Copy Status": f"{copy_status}"}
            # Unique key for item/machine pairing
            item_key = f"{item_program}{assigned_machine}"
            
            continue_check = read_json(checking_item_dict)
            bur_check = check_bur()
            if continue_check == "Continue Check":
                count += 1
                if email_sent == True or check_again == True:
                    print(item_program + " Checking Again")
                else:
                    print("NEW " + item_program)
                # Call to check_file function to check if folder path exist for machines
                check_assigned_machine()
                if copy_status != "":
                    read_json(checking_item_dict)
                    copy_status = ""
            elif continue_check == "Already Checked":
                pass
            elif continue_check == "Error":
                sys.exit()
        # Function to send email regarding items assigned to Walter machines
        def walter_email():
            global email_sent
            program_status = f"Program Status: Item Assigned To Walter Machine"
            walter_email_list.sort(key=lambda x: x[1])
            walter_items = "\n".join([f"{item} ➔ {machine}" for item, machine in walter_email_list])
            email_sent = False
            message = (f"Below item(s) have been assigned to the Walter machines.\nPlease update if needed and/or copy program to"
                    f" their respective machine(s).\n\nWalter Item(s):\n{walter_items}")
            send_email(message, recipient, cc_recipient, program_status)
        # Function to send email regarding programs that are missing .vgp
        def vgp_email():
            global email_sent
            program_status = f"Program Status: Program Missing .vgp"
            no_vgp_email_list.sort(key=lambda x: x[1])
            vgp_items = "\n".join([f"🔴 {item} ➔ {machine}\n" for item, machine in no_vgp_email_list])
            email_sent = False
            message = (f"Below item(s) are missing .vgp files.\nPlease update if needed and/or copy program to"
                    f" their respective machine(s).\n\nMissing .vgp Item(s):\n\n{vgp_items}")
            send_email(message, recipient, cc_recipient, program_status)
        def no_program_email():
            global email_sent
            program_status = f"Program Status: No Program Found"
            no_program_email_list.sort(key=lambda x: x[1])
            no_program_items = "\n".join([f"🔴 {item} ➔ {machine}  ({file})\n" for item, machine, file in no_program_email_list])
            email_sent = False
            message = (f"Below item(s) are missing their program.\nPlease update if needed and/or copy program to"
                    f" their respective machine(s).\n\nMissing Program(s):\n\n{no_program_items}")
            send_email(message, recipient, cc_recipient, program_status)
        if no_program_email_list != []:
            no_program_email()
        if no_vgp_email_list != []:
            vgp_email()
        if walter_email_list != []: 
            walter_email()
    if copied_program_list != []:
        print("\nCopied Program 📁")
        for item in copied_program_list:
            print(f"{item[0]} ➔  {item[1]} \u2713")
    print(f"\nChecked {count} Items")
    print("\nProgram Completed 🟢")
    print(run_time()) # Call to run_time()

def run_time():
    end_time =time.time()
    runtime = round(end_time - start_time, 5)
    if runtime < 60:
        runtime_txt = (f"Program runtime: {runtime} seconds")
    elif 60 <= runtime < 3600:
        runtime_txt = (f"Program runtime: {runtime/60} minutes")
    elif 3600 <= runtime < 86400:
        runtime_txt = (f"Program runtime: {runtime/3600} hours")
    elif 86400 <= runtime:
        runtime_txt = (f"Program runtime: {runtime/86400} days")
    return runtime_txt

# Function to check type of machine assigned
def check_assigned_machine():
    global get_machine_type, file_src, program_name_end, assigned_machine, item_program, copy_status
    program_name_end = ""
    # Adjust naming to get rid of coating on items
    # Handles standard burs
    if item_type == "Bur":
        program_name_end = bur_check
    # Handles coated naming
    if bur_check == "Not Bur" or bur_check == "Special Bur":
        if "/" in item_program:
            item_program = item_program.replace("/", "_")
        if "-" in item_program:
            item_name_split = item_program.split("-", 3)
            if item_program.startswith("DEN") and not item_program.startswith("DENMC"):
                program_name_end = item_name_split[0]
            else:
                if item_program.count("-") == 3:
                    program_name_end = item_name_split[0] + "-" + item_name_split[1] + "-" + item_name_split[2]
                else:
                    program_name_end = item_name_split[0] + "-" + item_name_split[1]
        else :
            program_name_end = item_program
    if testing == True:
                print(item_program + " ➔ " + program_name_end)
    get_machine_type = allowed_machines_dict.get(assigned_machine[0])
    # Used to distinguish the type of machine that is being checked
    if get_machine_type:
        global check_dst_date
        check_dst_date = ""
        # Rollomatic Machines
        if get_machine_type == "Rollomatic": 
            check_dst_date = 90 # 3 months compare rollomatic program to source
            file_dst = (f"{rollomatic_dst_path}\\{assigned_machine}")
            file_src = f"{rollomatic_src_path}\\{program_name_end}"
            check_file(file_src, file_dst)

        # Anca Machines
        elif get_machine_type == "Anca":
            check_dst_date = 180 # 6 months compare anca program to source
            machine_found = False
            if os.path.exists(anca_dst_path):
                for machines in os.listdir(anca_dst_path):
                    # Created a test folder inside A-18 only not on other machines - for testing
                    if testing == True:
                        assigned_machine = "A-22"
                    assigned_machine_split = assigned_machine.split("-",1)
                    assigned_machine_split = assigned_machine_split[0] + assigned_machine_split[1]
                    if machines.startswith(assigned_machine_split):
                        # Determine if machine uses TX7 or MX7
                        anca_machine_type = machines[-3:]
                        if anca_machine_type == "MX5" or anca_machine_type == "FX7":
                            anca_machine_type = "MX7"
                        if testing == False:
                            file_dst = (f"{anca_dst_path}\\{machines}\\tools")
                        else:
                            file_dst = (f"{anca_dst_path}\\{machines}\\tools\\Test1")
                        if os.path.exists(file_dst):
                            machine_found = True
                            break
                        else:
                            machine_found = False
                            break
                    else:
                        continue
                if machine_found == False:
                    program_status = f"Program Status: Anca Machine Not Found - {item_program}"
                    message = (f"{assigned_machine} machine was not found.\n\nPlease check {anca_dst_path}\n")
                    send_email(message, recipient, cc_recipient, program_status)
                # Check if program exist in source folder
                elif machine_found == True:
                    file_found = ""
                    if os.path.exists(anca_src_path):
                        for folder in os.listdir(anca_src_path):
                            if folder == "X-Do not use" or folder == "AI_RecycleBin" or folder == "Special":
                                continue
                            path = f"{anca_src_path}\\{folder}\\{program_name_end}.tom"
                            if os.path.exists(path):
                                if anca_machine_type == folder:
                                    file_src = path
                                    check_file_modification_date(file_src, file_dst)
                                    file_found = True
                                    break
                            else:
                                file_found = False
                        if file_found == False:
                            if email_sent == False:
                                no_program_email_list.append((item_program, assigned_machine, anca_machine_type))
                                copy_status = "Email Sent"
                            """program_status = f"Program Status: New Program Needed - {item_program}"
                            message = (f"Program file was not found in ({anca_machine_type}) folder.\n\nPlease create a new program for this item: {item_program} and copy program to {assigned_machine}.")
                            send_email(message, recipient, cc_recipient, program_status)"""
                    else:
                        program_status = f"Program Status: Anca Source Folder Missing"
                        message = f"Check {anca_src_path} file path."
                        send_email(message, recipient, cc_recipient, program_status)
            else:
                program_status = f"Program Status: Anca Main Folder Missing"
                message = f"Check Anca file path. Missing main Anca folder."
                send_email(message, recipient, cc_recipient, program_status)

        # Walter Machines
        elif get_machine_type == "Walter":
            if email_sent == False:
                walter_email_list.append((item_program, assigned_machine))
                copy_status = "Email Sent"
    else:
        program_status = f"Program Status: Unknown Machine Assigned Item"
        message = f"{item_program} assigned to {assigned_machine}\n\nMachine not recognized."
        send_email(message, recipient, cc_recipient, program_status)
# Check if source and destination files/folders exist
def check_file(file_src, file_dst,):

    src_exist = os.path.exists(file_src)
    dst_exist = os.path.exists(file_dst)
    # Check for machine existence
    if not dst_exist:
        if get_machine_type == "Rollomatic" or get_machine_type == "Anca":
            program_status = f"Program Status: {assigned_machine} Machine File Path Missing"
            message = (f"Check {assigned_machine} file path. Missing from main {get_machine_type} folder."
                    f"\nItem: {item_program}")
            send_email(message, recipient, cc_recipient, program_status)

    if dst_exist:
        # Check for program existence in machine
        file_dst = f"{file_dst}\\{item_program}"
        dst_exist = os.path.exists(file_dst)
        if src_exist:
            if not os.listdir(file_src):
                program_status = f"Program Status: {item_program} Folder Empty"
                message = (f"{item_program} - Program folder is empty please update and copy program to {assigned_machine}."
                        f"\nFile: {file_src}")
                send_email(message, recipient, cc_recipient, program_status)
            elif not dst_exist:
                os.mkdir(file_dst)
                check_file_modification_date(file_src, file_dst)
            elif src_exist and dst_exist:
                check_file_modification_date(file_src, file_dst)

        elif not src_exist:
            if email_sent == False:
                global copy_status
                location = file_src.split("\\")[-2:]
                location = "\\".join(location)
                no_program_email_list.append((item_program, assigned_machine, location))
                copy_status = "Email Sent"
            """program_status = f"Program Status: New Program Needed - {item_program}"
            message = (f"Program folder was not found.\n\nPlease create a new program for this item: {item_program} and copy program to {assigned_machine}."
                    f"\nFile: {file_src}")
            send_email(message, recipient, cc_recipient, program_status)"""

# Loops until all files are copied to destination
def copy_file(file_src, file_dst):
    global copy_status
    try:
        if get_machine_type == "Anca":
            shutil.copy2(file_src, file_dst)
            copy_status = "Success"
        if get_machine_type  == "Rollomatic":
            for file in os.listdir(file_src):
                copy_status = ""
                file_path = os.path.join(file_src, file)
                if os.path.isfile(file_path):
                    shutil.copy2(file_path, file_dst)
                    copy_status = "Success"
                else:
                    continue
        if copy_status == "Success":
            print(f"{item_program} copied to {assigned_machine}")
            copied_program_list.append((item_program, assigned_machine))

        log_result = f"{item_program} Files copied to {assigned_machine}"
        program_log(item_program, assigned_machine, log_result)
    # Handles the error if program is open
    except PermissionError as pe:
        if pe.winerror == 32:
            program_status = f"Program Status: Program Currently Open"
            message = f"{item_program} is currently running/open. Please move program to {assigned_machine} when program stops running."
            send_email(message, recipient, cc_recipient, program_status)
        
# Send email using smtp requires account information
def send_email(message, recipient, cc_recipient, program_status):
    if email_sent is True:
        return
    global copy_status
    copy_status = "Email Sent"
    try:
        global smtp_error
        smtp_server = config("SERVER", default = 'MAIL.COM')
        smtp_port = int(config("PORT", default = 25))
        sender_email = config("EMAIL_USER", default = 'Email')
        #sender_password = config("EMAIL_PASSWORD", default = 'Password') 
        
        email = EmailMessage()
        email["From"] = sender_email
        email["To"] = recipient
        email["CC"] = cc_recipient
        email["Subject"] = program_status
        if testing is True:
            email.set_content("TESTING PROGRAM IGNORE:\n\n" + message)
        else:
            email.set_content(message)
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            #server.starttls() # Use if not using internal relay
            #server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient, email.as_string())
        
        log_result = f"{program_status}--Email Sent"
        program_log(item_program, assigned_machine, log_result)
        return True
    except smtplib.SMTPAuthenticationError as sae:
        error_result = (f"SMTP Authentication Error\n{sae}")
        smtp_error = "SMTP Error"
        error_log(error_result)
    except smtplib.SMTPException as se:
        error_result = (f"SMTP Error\n{se}")
        smtp_error = "SMTP Error"
        error_log(error_result)
# Connects to local database
def connect_to_database():
    pass
# Check if the modification date is older than the check date
def check_file_modification_date(file_src, file_dst):
    global copy_status
    src_modification_date_list = []
    oldest_src_name = ""
    oldest_src_date = ""
    # Check source program folder for oldest file
    if get_machine_type == "Rollomatic":
        for file in os.listdir(file_src):
            if file.endswith(tuple(accepted_files_list)):
                file_path = os.path.join(file_src, file)
                if os.path.isfile(file_path):
                    src_modification_date_list.append((file, datetime.datetime.fromtimestamp(os.path.getmtime(file_path))))
        if src_modification_date_list == []:
            if email_sent == False:
                global copy_status
                no_vgp_email_list.append((item_program, assigned_machine))
                copy_status = "Email Sent"
            """program_status = "Program Status: No Program Found"

            message = (f"{item_program} - Program folder was not found with correct file type {accepted_files_list}."
                    f"\nAssigned machine: {assigned_machine}."
                    f"\nCheck if program has correct file type.")
            send_email(message, recipient, cc_recipient, program_status)"""

    if src_modification_date_list != [] or get_machine_type == "Anca":
        status = ""
        if get_machine_type == "Rollomatic":
            # Find oldest file date tuple pair from list
            oldest_file = min(src_modification_date_list, key=lambda x: x[1])
            oldest_src_name = oldest_file[0]
            oldest_src_date = oldest_file[1]
            if oldest_src_date > check_src_date:
                # Check destination folder for oldest file compared to source
                if os.path.exists(file_dst) and os.listdir(file_dst):
                    for program_src, program_src_date in src_modification_date_list:
                        dst_pro_path = os.path.join(file_dst, program_src)
                        if os.path.isfile(dst_pro_path):
                            program_dst_date = datetime.datetime.fromtimestamp(os.path.getmtime(dst_pro_path))
                            if program_src_date - program_dst_date > datetime.timedelta(days = check_dst_date):
                                copy_file(file_src, file_dst)
                            # Machine Program is up to date
                            else:
                                copy_status = "Up to Date"
                                break
                        else:
                            copy_file(file_src, file_dst)
                            break
                else:
                    copy_file(file_src, file_dst)
            else:
                status = "Old"
        # No folders to check just one file
        if get_machine_type == "Anca":
            oldest_src_name = os.path.basename(file_src)
            oldest_src_date = datetime.datetime.fromtimestamp(os.path.getmtime(file_src))
            if oldest_src_date > check_src_date:
                if os.path.exists(f"{file_dst}\\{oldest_src_name}"):
                    program_dst_date = datetime.datetime.fromtimestamp(os.path.getmtime(f"{file_dst}\\{oldest_src_name}"))
                    if oldest_src_date - program_dst_date > datetime.timedelta(days = check_dst_date):
                        copy_file(file_src, file_dst)
                    else:
                        copy_status = "Up to Date"
                else:
                    copy_file(file_src, file_dst)
            else:
                status = "Old"
        if status == "Old":
            program_status = "Program Status: Update Needed - File Outdated"
            message = (
                    f"Please Update The Following Program: {item_program}"
                    f"\n\nAssigned Machine: {get_machine_type} ({assigned_machine})"
                    f"\n\nOldest Program: {oldest_src_name}"
                    f"\nProgram Date: {oldest_src_date.strftime('%m-%d-%Y %I:%M:%S %p')}"
                    f"\n\n📁 {file_src}")
            send_email(message, recipient, cc_recipient, program_status)

# Program result log
def program_log(program_folder, machine_location, log_result):

    timestamp = datetime.datetime.now().strftime("%m-%d-%Y %I:%M:%S %p")
    runtime = run_time()
    # Writes to program log with lastest on top
    with open(program_log_path, "r") as f:
        program_log_content = f.read()
        f.close()

    with open(program_log_path, "w") as f:
        f.write(f"-----Program Executed at: {timestamp}-----\n{runtime}\nProgram: {program_folder}\nDestination Machine: {machine_location}\n{log_result}\n\n")
        f.write(program_log_content)
        f.close()
        
# Error log
def error_log(error_result):

    timestamp = datetime.datetime.now().strftime("%m-%d-%Y %I:%M:%S %p")
    
    # Writes to error log with lastest on top
    if not error_result == "":
        with open(error_log_path, "r") as f:
            error_log_content = f.read()
            f.close()
    if os.path.exists(error_log_path):
        with open(error_log_path, "w") as f:
            f.write(f"Error Executed at: {timestamp}\nError: {error_result}\n")
            f.write(f"\nFor item: {item_program} being assigned to: {assigned_machine}")
            f.write(error_log_content)
            f.close()
        if not smtp_error == "SMTP Error":
            program_status = "Error In Program Mover Program"
            message = (
                f"Error Executed at:\n\n{timestamp}\n{error_result}"
                f"\n\nFor item: {item_program} being assigned to: {assigned_machine}")
            recipient = error_recipient
            send_email(message, recipient, cc_recipient, program_status)
if __name__ == "__main__":  
    try:
        if testing is True or testing is False:
            main()
        else:
            raise ValueError("Testing variable must be set to True or False")
    except PermissionError as pe:
        error_result = f"{pe}"
        error_log(error_result)
    except Exception as e:
        error_result = f"{e}"
        error_log(error_result)
    finally:
        sys.exit()