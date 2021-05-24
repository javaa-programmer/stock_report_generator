import time
from masterdatafileupdater import MasterDataFileUpdater
from masterreportupdater import MasterReportUpdater
import stockreportgeneratorhelper as srgh
from dailyreportgenerator import DailyReportGenerator
from zipfile import ZipFile
import shutil
import os
from jproperties import Properties
from staticconfiguration import StaticConfiguration


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
    temp_file_name = sc.temp_file_path + 'PR' + current_date + '.zip'
    with ZipFile(temp_file_name, 'r') as zipObj:
        zipObj.extractall(sc.temp_file_path)

    file_name = sc.temp_file_path + 'Pd' + current_date + '.csv'
    shutil.copy(file_name, sc.input_file_path)

    for filename in os.listdir(sc.temp_file_path):
        file_path = os.path.join(sc.temp_file_path, filename)
        try:
            shutil.rmtree(file_path)
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

# Check if the run date is holiday or not
is_holiday = srgh.check_holiday(srgh.create_date(current_date), sc)

if not is_holiday:

    # Unzip the file, copy to input directory
    # and delete the unused files.
    prepare_file()

    # Set the flags and Unzip the file, copy to input directory
    # and delete the unused files.
    srgh.set_app_process_flags(srgh.create_date(current_date), sc)

    # Update the master data excel
    mdfu = MasterDataFileUpdater(sc.master_data_file_name, 'Details', current_date, sc)
    mdfu.update_master_data()

    # update the master report
    mru = MasterReportUpdater(sc.master_data_file_name, 'Details', current_date, sc)
    mru.update_master_report()

    # Generate Daily Reports
    drg = DailyReportGenerator(sc.master_data_file_name, 'Details', current_date, sc)
    drg.generate_daily_reports()

    # Generate Weekly Reports
    # generate_weekly_reports()

    # Generate Monthly Reports
    # generate_monthly_reports()

else:
    print(f'{srgh.create_date(current_date).date()} is Holiday and Market is Closed')

print(f'Time taken to complete the process : {time.time() - start_time}')
