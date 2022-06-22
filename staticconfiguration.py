
class StaticConfiguration:

    def __init__(self, configs):
        self.input_file_path = configs.get("input_file_path").data
        self.output_file_path = configs.get("output_file_path").data
        self.other_file_path = configs.get("other_file_path").data
        self.master_data_file_name = self.output_file_path + configs.get("master_data_file_name").data
        self.master_report_name = self.output_file_path + configs.get("master_report_name").data
        self.bulk_file_name = self.input_file_path + configs.get("bulk_file_name").data
        self.block_file_name = self.input_file_path + configs.get("block_file_name").data
        self.master_scrip_list = self.other_file_path + configs.get("master_scrip_list").data
        self.nse_holiday_list = self.other_file_path + configs.get("nse_holiday_list").data
        self.App_Process_Calendar_file = self.other_file_path + configs.get("App_Process_Calendar_file").data
        self.daily_report_name = configs.get("daily_report_name").data
        self.sheet_name_price_volume = configs.get("sheet_name_price_volume").data
