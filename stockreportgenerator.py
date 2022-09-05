import time
from masterdatafileupdater import MasterDataFileUpdater
from masterreportupdater import MasterReportUpdater
import stockreportgeneratorhelper as srgh
from dailyreportgenerator import DailyReportGenerator
from zipfile import ZipFile
import os
from jproperties import Properties
from staticconfiguration import StaticConfiguration
from datetime import timedelta


# Generates the weekly reports
# Will be implemented later and moved to other class
def generate_weekly_reports():
    pass


# Generates the monthly reports
# Will be implemented later and moved to other class
def generate_monthly_reports():
    pass


# Unzip the file, copy to input directory
# and delete the unused files.
def prepare_file():
    temp_file_name = sc.input_file_path + 'PR' + current_date + '.zip'
    try:
        with ZipFile(temp_file_name, 'r') as zipObj:
            zipObj.extractall(sc.input_file_path)
    except FileNotFoundError:
        return False

    return True


def delete_files():
    for filename in os.listdir(sc.input_file_path):
        file_path = os.path.join(sc.input_file_path, filename)
        try:
            os.remove(file_path)
        except OSError:
            os.remove(file_path)


configs = Properties()


def load_config():
    with open("D:\\personal\\stock-market\\others\\application_config.properties", 'rb') as config_file:
        configs.load(config_file)
    sc = StaticConfiguration(configs)
    return sc


start_time = time.time()

current_date = srgh.get_current_date()
sc = load_config()
from_date = srgh.create_date(current_date) - timedelta(days=45)

# Check if the run date is holiday or not
is_holiday = srgh.check_holiday(srgh.create_date(current_date), sc)

if not is_holiday:

    # Unzip the file, copy to input directory
    # and delete the unused files.
    file_available = prepare_file()

    if file_available:

        data_start_time = time.time()
        # Update the master data excel
        mdfu = MasterDataFileUpdater(sc.master_data_file_name, 'Details', current_date, sc)
        mdfu.update_master_data()
        print(f'Data file update complete. Time taken {time.time() - data_start_time} secs')

        # update the master report
        mr_start_time = time.time()
        mru = MasterReportUpdater(sc.master_data_file_name, 'Details', current_date, from_date.date(), sc)
        mru.update_master_report()
        print(f'Master report update complete. Time taken {time.time() - mr_start_time} secs')

        # Generate Daily Reports
        dr_start_time = time.time()
        drg = DailyReportGenerator(sc.master_data_file_name, 'Details', current_date, sc)
        drg.generate_daily_reports()
        print(f'Daily report generation is complete. Time taken {time.time() - dr_start_time} secs')

        # Generate Weekly Reports
        # generate_weekly_reports()

        # Generate Monthly Reports
        # generate_monthly_reports()
    else:
        print("File not found for Generating Report")

else:
    print(f'{srgh.create_date(current_date).date()} + is Holiday and Market is Closed')

delete_files()

print(f'Time taken to complete the process : {time.time() - start_time}')
current_date = input("Press any key to continue... ")
