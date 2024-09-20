import os
import json
from workalendar.europe import UnitedKingdom, Ireland
from datetime import datetime, date
import logging
 
BASE_PATH = r"Logging File Path"
 
def setup_logging(base_path):
   
    now = datetime.now()
    folder_name = now.strftime("%m%Y")
    folder_path = os.path.join(base_path, folder_name)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    log_file_name = now.strftime("%d%m.log")
    log_file_path = os.path.join(folder_path, log_file_name)
    logging.basicConfig(
        filename=log_file_path,
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
 
    logging.debug("Logging setup complete.")
setup_logging(BASE_PATH)

def get_Enfusion_date():
    cal1 = UnitedKingdom()
    cal2 = Ireland()
    UK_Previous_Bus_Day = cal1.add_working_days(date.today(), -1)
    Ireland_Previous_Bus_Day = cal2.add_working_days(date.today(), -1)
    if UK_Previous_Bus_Day == Ireland_Previous_Bus_Day or UK_Previous_Bus_Day < Ireland_Previous_Bus_Day:
        Enfusion_date = UK_Previous_Bus_Day
    else:
        Enfusion_date = Ireland_Previous_Bus_Day
    return Enfusion_date

with open("Dictionary File Path") as f:
    data = json.load(f)

NT_Downloads = {'API URL part 2' : { "Date" : str(get_Enfusion_date()),
                                                  "FILE_NAME": "NT_NAV"}
} 
# If API predicted available add into above dictionary 

script = "NT_API_code_1.py"
pythonexe = "python3"
report_name1 = "--api_url"
date1 = "--date"
nt_fund_code1 = "--nt_fund_code"
file_path1 = "--file_path"
file_path2 = "Directory to download to"
file_name1 = "--file_name"

Existing_NT_NAV_Files = os.listdir("Directory for previously downloaded files")

os.chdir("Directory path for API grab script")
import subprocess
for i in NT_Downloads:
    report_name2 = 'API URL part 1' + i
    date2 = NT_Downloads[i]["Date"]
    logging.info(date2)
    for x in data:
        file_name2 = x + "_" + NT_Downloads[i]["FILE_NAME"] + "_" + str(date.today()) + ".xlsx"
        logging.info(file_name2)
        if file_name2 in Existing_NT_NAV_Files:
            continue
        else:
            nt_fund_code2 = data[x]["NT_FUND_CODE"]
            result = subprocess.run([pythonexe, script, report_name1, report_name2, nt_fund_code1, nt_fund_code2, date1, date2, file_path1, file_path2, file_name1, file_name2], capture_output=True, text=True)
            logging.debug(result)