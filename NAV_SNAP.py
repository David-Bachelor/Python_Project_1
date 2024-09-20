import pandas as pd
import datetime
import os
from io import BytesIO
import msoffcrypto
import logging
import json
from workalendar.europe import UnitedKingdom, Ireland 
from datetime import date, datetime

# Setup logging

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


# Load JSON data
with open('Dictionary File Path', 'r') as f:
    data = json.load(f)

# Initialize results dictionary
Results = {}

# Get business days
cal = UnitedKingdom()

# Download Enfusion Files
def Run_Enfusion_Downloader():
    with open('Filepath\\Call_Enfusion_Grab_Final.py') as file:
        exec(file.read())

Run_Enfusion_Downloader()

# Download NT Files
def Run_NT_Downloader():
    with open('Filepath\\API_code_caller.py') as file:
        exec(file.read())

Run_NT_Downloader()

def get_prediction_date():
    cal1 = UnitedKingdom()
    cal2 = Ireland()
    UK_Previous_Bus_Day = cal1.add_working_days(date.today(), -1)
    Ireland_Previous_Bus_Day = cal2.add_working_days(date.today(), -1)
    if UK_Previous_Bus_Day == Ireland_Previous_Bus_Day or UK_Previous_Bus_Day < Ireland_Previous_Bus_Day:
        prediction_date = UK_Previous_Bus_Day
    else:
        prediction_date = Ireland_Previous_Bus_Day
    logging.info('Prediction/Enfusion Date:', prediction_date)
    return prediction_date

def get_filtered_files(path, filters, exclude_ext=None):
    files = os.listdir(path)
    filtered_files = [file for file in files if all(f in file for f in filters)]
    if exclude_ext:
        filtered_files = [file for file in filtered_files if not file.endswith(exclude_ext)]
    return filtered_files

def get_file_by_date(files, date_str):
    return [file for file in files if date_str in file]

def decrypt_file(file_path, password):
    with open(file_path, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        decrypted_file = BytesIO()
        office_file.decrypt(decrypted_file)
    decrypted_file.seek(0)
    return decrypted_file

def calculate_nav_difference(nav_df, pred_df, enfusion_df, prediction_row):
    logging.info(nav_df)
    nt_nav = nav_df['ColumnLabel'][nav_df['ColumnLabel'] == 'Rowname'].dropna().values[0]
    column_name = pred_df.columns[pred_df.isin(['CellValue']).any()].tolist()[0]
    subs_reds = pred_df[column_name][pred_df['Summary Flows by Fund / Share Class'] == prediction_row].dropna().values[0]
    enfusion_nav = enfusion_df['Column_Label'].sum()
    difference = (enfusion_nav - subs_reds) - nt_nav
    pct_difference = (difference / nt_nav) * 100
    basis_points = pct_difference * 100
    logging.info(nt_nav, enfusion_nav, subs_reds, difference, pct_difference, basis_points)
    return nt_nav, enfusion_nav, subs_reds, difference, pct_difference, basis_points

def handle_missing_files(i, prediction_files, enfusion_files, nav_files):
    Error_message = []
    if not prediction_files:
        Error_message.append(f"{i} Predicted SUBS/REDS File Missing")
        logging.error(f"{i} Predicted SUBS/REDS File Missing")
    if not enfusion_files:
        Error_message.append(f"{i} Enfusion file Missing")
        logging.error(f"{i} Enfusion file Missing")
    if not nav_files:
        logging.info(f"{i} NT NAV file Missing")
    if Error_message == []:
        Error_message = ""
    results = {
        'NT NAV': 'NA',
        'ENFUSION NAV': 'NA',
        'SUBS/REDS': 'NA', 
        'Difference': 'NA', 
        '% Difference': 'NA', 
        'Basis Points': 'NA', 
        'Errors': Error_message
        }
    Results[i] = results

def process_files_for_fund(data, i):
    base_nt_api_path = r"Filepath"
    base_predicted_files_path = r"Filepath"
    Enfusion_path = r"Filepath"
    fund_code = data[i]['FUND_NT_FILECODE']
    prediction_code = data[i]['FUND_PREDICTION_FILECODE']
    prediction_row = data[i]['FUND_PREDICTION_ROWNAME']
    password = "Password"
    Error_message = []

    nav_files = get_filtered_files(base_nt_api_path, [i, 'NT_NAV'])
    nav_date = date.today()
    nav_files = get_file_by_date(nav_files, str(nav_date))
    
    prediction_date = get_prediction_date()
    prediction_date_str = prediction_date.strftime("%Y%m%d")
    prediction_files = get_file_by_date(get_filtered_files(base_predicted_files_path, [prediction_code]), prediction_date_str)
    logging.info(prediction_files)

    enfusion_files = get_filtered_files(Enfusion_path, ['Daily Nav Check', i])
    logging.info(enfusion_files)

    if not nav_files or not prediction_files or not enfusion_files:
        handle_missing_files(i, prediction_files, enfusion_files, nav_files)
        return Error_message

    nav_file = nav_files[0]
    logging.info(nav_file)
    nav_df = pd.read_excel(os.path.join(base_nt_api_path, nav_file))

    prediction_file = prediction_files[0]
    logging.info(prediction_file)
    decrypted_file = decrypt_file(os.path.join(base_predicted_files_path, prediction_file), password)
    pred_df = pd.read_excel(decrypted_file, 'Summary Flows by Fund or Class')

    enfusion_file = enfusion_files[0]
    logging.info(enfusion_file)
    enfusion_df = pd.read_csv(os.path.join(Enfusion_path, enfusion_file))

    nt_nav, enfusion_nav, subs_reds, difference, pct_difference, basis_points = calculate_nav_difference(nav_df, pred_df, enfusion_df, prediction_row)

    Results[i] = {
        'NT NAV': nt_nav,
        'ENFUSION NAV': enfusion_nav,
        'SUBS/REDS': subs_reds,
        'Difference': difference,
        '% Difference': pct_difference,
        'Basis Points': basis_points,
        'Errors': ""
    }
    logging.info(Results)
for i in data:
    process_files_for_fund(data, i)

nav_table = pd.DataFrame.from_dict(Results, orient='index')


logging.info(nav_table)

def Write_To_XLSX(nav_table):
    file_name = "C:\\Users\\David.Bachelor\\Blank_Test.xlsx"
    nav_table.to_excel(file_name)


Write_To_XLSX(nav_table)


