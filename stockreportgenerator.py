import time
from masterdatafileupdater import MasterDataFileUpdater
from masterreportupdater import MasterReportUpdater
import stockreportgeneratorhelper as srgh
from dailyreportgenerator import DailyReportGenerator
import directorypaths as dp
from zipfile import ZipFile
import shutil
import os


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

    temp_file_name = dp.temp_file_path + 'PR' + srgh.current_date_str + '.zip'
    with ZipFile(temp_file_name, 'r') as zipObj:
        zipObj.extractall(dp.temp_file_path)

    file_name = dp.temp_file_path + 'Pd' + srgh.get_current_date() + '.csv'
    shutil.copy(file_name, dp.input_file_path)

    for filename in os.listdir(dp.temp_file_path):
        file_path = os.path.join(dp.temp_file_path, filename)
        try:
            shutil.rmtree(file_path)
        except OSError:
            os.remove(file_path)


start_time = time.time()

# Check if the run date is holiday or not
is_holiday = srgh.check_holiday(srgh.gl_formatted_date)

if not is_holiday:

    # Unzip the file, copy to input directory
    # and delete the unused files.
    prepare_file()

    # Set the flags and Unzip the file, copy to input directory
    # and delete the unused files.
    srgh.set_app_process_flags(srgh.gl_formatted_date)

    # Update the master data excel
    mdfu = MasterDataFileUpdater(dp.master_data_file_name, 'Details', srgh.current_date_str)
    mdfu.update_master_data()

    # update the master report
    mru = MasterReportUpdater(dp.master_data_file_name, 'Details', srgh.current_date_str)
    mru.update_master_report()

    # Generate Daily Reports
    drg = DailyReportGenerator(dp.master_data_file_name, 'Details', srgh.current_date_str)
    drg.generate_daily_reports()

    # Generate Weekly Reports
    # generate_weekly_reports()

    # Generate Monthly Reports
    # generate_monthly_reports()

else:
    print(f'{srgh.gl_formatted_date.date()} is Holiday and Market is Closed')

print(f'Time taken to complete the process : {time.time() - start_time}')
