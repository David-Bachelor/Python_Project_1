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
 
def Get_Username():
    user = os.getlogin()
    print(user)
    if user == "User":
        username = user + "@Domain-name.com"
        return username
    else:
        username = user + "@Domain-name.com"
        return username

with open(r"Fund Dictionary Filepath") as f:
    data = json.load(f)
 
report_names=[]
 
for i in data:
    report_name = data[i]['ENFUSION']
    report_names.append(report_name)
 
username1 = "--username"
username2 = Get_Username()

password1 = "--password"
password2 = input("Enter Enfusion Password:")

report_names = report_names
report_name1 = '--report_name'
date1 = "--date"
date2 = str(get_Enfusion_date())
logging.info(date2)
file_path1 = "--file_path"
file_path2 = r"Filepath\For\Enfusion_Download"

script = 'Enfusiongrab.py'
pythonexe = 'python3'
 
os.chdir("C:\\Users\\David.Bachelor\\NAV_code_Final")
import subprocess
for i in report_names:
    report_name = 'shared/Daily Nav Check Reports/' + i + '.ppr'
    logging.info(report_name)
    file_name = "--file_name"
    file_name2 = i + ".csv"
    result = subprocess.run([pythonexe, script, username1, username2, password1, password2, report_name1, report_name, date1, date2, file_path1, file_path2, file_name, file_name2], capture_output=True, text=True)
    logging.debug(result)