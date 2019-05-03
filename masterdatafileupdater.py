import pandas as pd
import directorypaths as dp
import stockreportgeneratorhelper as srgh


class MasterDataFileUpdater:

    # Constructor
    # file_name: The master data file name
    # sheet_name: The active sheet name
    # current_date: Report generation date as String
    def __init__(self, file_name, sheet_name, current_date):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.current_date = current_date

    # Update the master data excel
    def update_master_data(self):
        # Create the input file name
        input_file_name = dp.input_file_path + 'Pd' + self.current_date + '.csv'

        # Read the Price File
        scrip_details = pd.read_csv(input_file_name)

        nifty_details = scrip_details[scrip_details['SECURITY'] == 'Nifty 50'].copy(deep=True)
        nifty_details.rename(columns={'SECURITY': 'NAME'}, inplace=True)
        nifty_details['SYMBOL'] = 'NIFTY'
        nifty_details = nifty_details[
            ['SYMBOL', 'NAME', 'SERIES', 'HI_52_WK', 'LO_52_WK', 'PREV_CL_PR', 'OPEN_PRICE', 'HIGH_PRICE', 'LOW_PRICE',
            'CLOSE_PRICE', 'NET_TRDQTY']]

        # Read the Scrip List
        scrip_list = pd.read_excel(dp.master_scrip_list, "SCRIP_LIST")

        # Prepare the list for Selected List
        selected_list = pd.merge(scrip_list, scrip_details, left_on=["SYMBOL", "SERIES"], right_on=["SYMBOL", "SERIES"])

        # Fetch the fields to be updated in master file
        selected_fields = selected_list[['SYMBOL', 'NAME', 'SERIES', 'HI_52_WK', 'LO_52_WK', 'PREV_CL_PR', 'OPEN_PRICE',
                                     'HIGH_PRICE', 'LOW_PRICE', 'CLOSE_PRICE', 'NET_TRDQTY']]

        nifty_details = pd.concat([nifty_details, selected_fields], ignore_index=True)

        # Add the Date Column in Data Frame
        nifty_details.insert(3, "TRADE_DATE", [srgh.gl_formatted_date.date()]*len(nifty_details.index))

        # Append the data in xlsx file
        srgh.append_df_to_excel(self.file_name, nifty_details, self.sheet_name)
