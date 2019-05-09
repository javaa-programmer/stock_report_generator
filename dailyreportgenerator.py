import pandas as pd
import stockreportgeneratorhelper as srgh
from openpyxl.styles import PatternFill
import directorypaths as dp
from openpyxl import load_workbook
import numpy as np
from datetime import timedelta

class DailyReportGenerator:

    def __init__(self, input_file_name, data_sheet_name, current_date_str):
        self.input_file_name = input_file_name
        self.data_sheet_name = data_sheet_name
        self.current_date_str = current_date_str

    # Generate Daily Reports
    # Generate the Price Volume Report
    def generate_daily_reports(self):
        current_date = srgh.create_date(self.current_date_str)
        previous_date = srgh.offset_business_day(current_date, 1)
        report_name = dp.output_file_path + dp.daily_report_name + '_' + str(current_date.date()) + '.xlsx'

        DailyReportGenerator.generate_price_volume_report = staticmethod(
                                   DailyReportGenerator.generate_price_volume_report)
        DailyReportGenerator.generate_price_volume_report(self, current_date, previous_date, report_name)

        DailyReportGenerator.generate_new_52_week_high_low_report = staticmethod(
                               DailyReportGenerator.generate_new_52_week_high_low_report)
        DailyReportGenerator.generate_new_52_week_high_low_report(self, current_date, report_name)

        DailyReportGenerator.generate_volatile_stock_day = staticmethod(DailyReportGenerator.generate_volatile_stock_day)
        DailyReportGenerator.generate_volatile_stock_day(self, current_date, report_name)

        DailyReportGenerator.generate_trending_scrip_list = staticmethod(DailyReportGenerator.generate_trending_scrip_list)
        DailyReportGenerator.generate_trending_scrip_list(self, current_date, report_name)

    # Generate the report for the shares whose close price is increased
    # or decreased three consecutive days.
    def generate_trending_scrip_list(self, current_date, report_name):
        master_data = pd.read_excel(self.input_file_name, self.data_sheet_name)

        to_date = srgh.create_date(srgh.current_date_str)
        from_date = srgh.offset_business_day(current_date, 2)

        increased_price_data = master_data[(master_data['TRADE_DATE'] <= to_date) &
                                  (master_data['TRADE_DATE'] >= from_date) &
                                  (master_data['CLOSE_PRICE'] >= master_data['PREV_CL_PR'])]

        increased_price_data['freq'] = increased_price_data.groupby('SYMBOL')['SYMBOL'].transform('count').copy(deep=True)
        increased_price_data = increased_price_data[(increased_price_data['freq'] == 3)].copy(deep=True)

        increased_price_data = increased_price_data[['SYMBOL','NAME', 'TRADE_DATE', 'PREV_CL_PR', 'CLOSE_PRICE', 'NET_TRDQTY']].copy(deep=True)
        temp_df = increased_price_data[(increased_price_data['TRADE_DATE'] == from_date)]

        decreased_price_data = master_data[(master_data['TRADE_DATE'] <= to_date) &
                                           (master_data['TRADE_DATE'] >= from_date) &
                                           (master_data['CLOSE_PRICE'] <= master_data['PREV_CL_PR'])]

        decreased_price_data['freq_dr'] = decreased_price_data.groupby('SYMBOL')['SYMBOL'].transform('count').copy(
            deep=True)
        decreased_price_data = decreased_price_data[(decreased_price_data['freq_dr'] == 3)].copy(deep=True)

        decreased_price_data = decreased_price_data[
            ['SYMBOL', 'NAME', 'TRADE_DATE', 'PREV_CL_PR', 'CLOSE_PRICE', 'NET_TRDQTY']].copy(deep=True)
        temp_dr_df = decreased_price_data[(decreased_price_data['TRADE_DATE'] == from_date)]

        app_date = from_date
        while app_date < to_date:
            app_date = app_date + timedelta(days=1)
            while srgh.check_holiday(app_date):
                app_date = app_date + timedelta(days=1)

            temp_df1 = increased_price_data[(increased_price_data['TRADE_DATE'] == app_date)]
            temp_df1 = temp_df1[['SYMBOL','NAME', 'CLOSE_PRICE', 'NET_TRDQTY']]

            temp_dr_df1 = decreased_price_data[(decreased_price_data['TRADE_DATE'] == app_date)]
            temp_dr_df1 = temp_dr_df1[['SYMBOL', 'NAME', 'CLOSE_PRICE', 'NET_TRDQTY']]

            try:
                temp_df = pd.merge(temp_df, temp_df1, left_on = ['SYMBOL','NAME'], right_on = ['SYMBOL','NAME'])
                temp_df.rename(columns=srgh.cons_increased_header1, inplace=True)

                temp_dr_df = pd.merge(temp_dr_df, temp_dr_df1, left_on=['SYMBOL', 'NAME'], right_on=['SYMBOL', 'NAME'])
                temp_dr_df.rename(columns=srgh.cons_increased_header1, inplace=True)

            except IndexError:
                temp_df = temp_df if not temp_df.empty else temp_df1
                temp_dr_df = temp_dr_df if not temp_dr_df.empty else temp_dr_df1

            while srgh.check_holiday(app_date):
                app_date = app_date + timedelta(days=1)
        temp_df.drop(columns=['TRADE_DATE'], inplace=True)
        temp_dr_df.drop(columns=['TRADE_DATE'], inplace=True)

        book = load_workbook(report_name)
        writer = pd.ExcelWriter(report_name, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        temp_df.to_excel(writer, "Trendies Technical - I", startrow=2, index=False)
        writer.save()
        DailyReportGenerator.format_cons_increase_report = staticmethod(
            DailyReportGenerator.format_cons_increase_report)
        DailyReportGenerator.format_cons_increase_report(self, report_name)

        book = load_workbook(report_name)
        sheet = book['Trendies Technical - I']

        DailyReportGenerator.update_decr_scrip_list = staticmethod(
            DailyReportGenerator.update_decr_scrip_list)
        DailyReportGenerator.update_decr_scrip_list(self,temp_dr_df, sheet.max_row + 3, report_name)

        DailyReportGenerator.format_cons_increase_report = staticmethod(
            DailyReportGenerator.format_cons_decrease_report)
        DailyReportGenerator.format_cons_decrease_report(self, sheet.max_row + 2, report_name)

    # Update
    def update_decr_scrip_list(self, decreased_scrip_df, star_row, report_name):
        book = load_workbook(report_name)
        writer = pd.ExcelWriter(report_name, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        # selected_list.to_excel(writer, sheet_name, header=None, index=False)

        decreased_scrip_df.to_excel(writer, 'Trendies Technical - I',startrow=star_row, index=False)
        writer.save()

    # Format Cons Increase Report
    def format_cons_increase_report(self, report_name):
        to_date = srgh.create_date(srgh.current_date_str)
        from_date = srgh.offset_business_day(to_date, 2)

        book = load_workbook(report_name)
        sheet = book['Trendies Technical - I']
        sheet.merge_cells('A1:I1')
        sheet.cell(row=1, column=1).value = 'Scrips with Price Increased Three Consecutive Session'
        sheet.merge_cells('A2:B2')
        sheet.cell(row=2, column=1).value = 'Scrip Details'
        max_column = sheet.max_column
        curr_column = 4
        while curr_column < max_column:
            sheet.merge_cells(start_row=2, start_column=curr_column, end_row=2, end_column=curr_column + 1)

            while srgh.check_holiday(from_date):
                from_date = from_date + timedelta(days=1)

            sheet.cell(row=2, column=curr_column).value = from_date.date()
            curr_column = curr_column + 2
            from_date = from_date + timedelta(days=1)

        curr_column = 1
        max_rows = sheet.max_row
        curr_row = 1
        while curr_row <= max_rows:
            while curr_column <= max_column:
                sheet.cell(curr_row, curr_column).border = srgh.thin_border
                if curr_row == 3 or curr_row == 2:
                    sheet.cell(curr_row, curr_column).font = srgh.font_header
                    sheet.cell(curr_row, curr_column).fill = PatternFill(start_color='D3D3D3', fill_type="solid")
                    sheet.cell(curr_row, curr_column).alignment = srgh.align_header
                elif curr_row == 1:
                    sheet.cell(curr_row, curr_column).font = srgh.font_header
                    sheet.cell(curr_row, curr_column).fill = PatternFill(start_color='CAFF33', fill_type="solid")
                    sheet.cell(curr_row, curr_column).alignment = srgh.align_header
                else:
                    sheet.cell(curr_row, curr_column).font = srgh.font_body
                    if curr_column < 3:
                        sheet.cell(curr_row, curr_column).alignment = srgh.align_body_str
                    else:
                        sheet.cell(curr_row, curr_column).alignment = srgh.align_body_num
                curr_column = curr_column + 1
            sheet.row_dimensions[curr_row].height = 20  # In pixels
            curr_row = curr_row + 1
            curr_column = 1

        curr_row_no = 4
        for rows in sheet.iter_rows(min_row=4, max_row=max_rows, min_col=1):
            for cell in rows:
                if curr_row_no % 2 == 1:
                    cell.fill = PatternFill(start_color="f7ec8f", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="edead7", fill_type="solid")
            curr_row_no = curr_row_no + 1
        sheet.column_dimensions['A'].width = 18
        sheet.column_dimensions['B'].width = 25
        sheet.column_dimensions['C'].width = 22
        sheet.column_dimensions['D'].width = 15
        sheet.column_dimensions['E'].width = 15
        sheet.column_dimensions['F'].width = 15
        sheet.column_dimensions['G'].width = 15
        sheet.column_dimensions['H'].width = 15
        sheet.column_dimensions['I'].width = 15

        book.save(report_name)

    # Format Cons Increase Report
    def format_cons_decrease_report(self, start_row, report_name):
        to_date = srgh.create_date(srgh.current_date_str)
        from_date = srgh.offset_business_day(to_date, 2)

        book = load_workbook(report_name)

        sheet = book['Trendies Technical - I']
        cell_range = 'A' + str(start_row) + ':'+'I'+str(start_row)
        sheet.merge_cells(cell_range)
        sheet.cell(row=start_row, column=1).value = 'Scrips with Price Decreased Three Consecutive Session'
        cell_range = 'A' + str(start_row + 1) + ':'+'B'+str(start_row + 1)
        sheet.merge_cells(cell_range)
        sheet.cell(row=start_row + 1, column=1).value = 'Scrip Details'
        max_column = sheet.max_column
        curr_column = 4
        while curr_column < max_column:
            sheet.merge_cells(start_row=start_row + 1, start_column=curr_column, end_row=start_row + 1, end_column=curr_column + 1)

            while srgh.check_holiday(from_date):
                from_date = from_date + timedelta(days=1)

            sheet.cell(row=start_row +1, column=curr_column).value = from_date.date()
            curr_column = curr_column + 2
            from_date = from_date + timedelta(days=1)

        curr_column = 1
        max_rows = sheet.max_row
        curr_row = start_row
        while curr_row <= max_rows:
            while curr_column <= max_column:
                sheet.cell(curr_row, curr_column).border = srgh.thin_border
                if curr_row == start_row + 1  or curr_row == start_row + 2:
                    sheet.cell(curr_row, curr_column).font = srgh.font_header
                    sheet.cell(curr_row, curr_column).fill = PatternFill(start_color='D3D3D3', fill_type="solid")
                    sheet.cell(curr_row, curr_column).alignment = srgh.align_header
                elif curr_row == start_row:
                    sheet.cell(curr_row, curr_column).font = srgh.font_header
                    sheet.cell(curr_row, curr_column).fill = PatternFill(start_color='F6646B', fill_type="solid")
                    sheet.cell(curr_row, curr_column).alignment = srgh.align_header
                else:
                    sheet.cell(curr_row, curr_column).font = srgh.font_body
                    if curr_column < 3:
                        sheet.cell(curr_row, curr_column).alignment = srgh.align_body_str
                    else:
                        sheet.cell(curr_row, curr_column).alignment = srgh.align_body_num
                curr_column = curr_column + 1
            sheet.row_dimensions[curr_row].height = 20  # In pixels
            curr_row = curr_row + 1
            curr_column = 1

        curr_row_no = start_row + 3
        for rows in sheet.iter_rows(min_row=start_row + 3, max_row=max_rows, min_col=1):
            for cell in rows:
                if curr_row_no % 2 == 1:
                    cell.fill = PatternFill(start_color="f7ec8f", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="edead7", fill_type="solid")
            curr_row_no = curr_row_no + 1
        sheet.column_dimensions['B'].width = 25
        sheet.column_dimensions['A'].width = 12

        book.save(report_name)

    # Generate the Price Volume Report
    def generate_price_volume_report(self, current_date, previous_date, report_name):
        master_data = pd.read_excel(self.input_file_name, self.data_sheet_name)
        master_data = master_data[(master_data['TRADE_DATE'] <= current_date) &
                                  (master_data['TRADE_DATE'] >= previous_date)]

        master_data = master_data[['SYMBOL', 'NAME', 'TRADE_DATE', 'PREV_CL_PR', 'CLOSE_PRICE', 'NET_TRDQTY']]

        # Copy the previous day's volume in current days data
        master_data['PREV_VOL'] = master_data.groupby(['SYMBOL'])['NET_TRDQTY'].shift(1)

        # Filter the current date records
        master_data = master_data[(master_data['TRADE_DATE'] == current_date)]

        change = (master_data[master_data.columns[4]] - master_data[master_data.columns[3]])
        change_per = ((change * 100) / master_data[master_data.columns[3]])

        vol_change = (master_data[master_data.columns[5]] - master_data[master_data.columns[6]])
        vol_change_per = ((vol_change * 100) / master_data[master_data.columns[6]])

        master_data.insert(5, "Change", change)
        master_data.insert(6, "Change(%)", change_per)

        master_data.insert(7, "Volume Change", vol_change)
        master_data.insert(8, "Volume Change(%)", vol_change_per)

        master_data.drop(columns=['TRADE_DATE'], inplace=True)

        master_data.sort_values(["Change(%)"], axis=0, ascending=False, inplace=True)

        price_incr_vol_incr = master_data[(master_data['Change'] >= 0) &
                                          (master_data['Volume Change'] >= 0)]

        price_incr_vol_decr = master_data[(master_data['Change'] >= 0) &
                                          (master_data['Volume Change'] <= 0)]

        price_decr_vol_incr = master_data[(master_data['Change'] <= 0) &
                                          (master_data['Volume Change'] >= 0)]

        price_decr_vol_decr = master_data[(master_data['Change'] <= 0) &
                                          (master_data['Volume Change'] <= 0)]

        narration1 = pd.DataFrame({'SYMBOL': 'Price Increased - Volume Increased', 'NAME': ' ',
                                   'PREV_CL_PR': '', 'CLOSE_PRICE': '', 'Change': '', 'Change(%)': '',
                                   'Volume Change': '', 'Volume Change(%)': '', 'NET_TRDQTY': '', 'PREV_VOL': ''},
                                  index=[0])

        cust_header = pd.DataFrame({'SYMBOL': 'SYMBOL', 'NAME': 'Name', 'PREV_CL_PR': 'Previous Close',
                                    'CLOSE_PRICE': 'Close Price', 'Change': 'Change', 'Change(%)': 'Change(%)',
                                    'Volume Change': 'Volume Change', 'Volume Change(%)': 'Volume Change(%)',
                                    'NET_TRDQTY': 'Volume', 'PREV_VOL': 'Prev. Volume'}, index=[0])

        narration2 = pd.DataFrame({'SYMBOL': 'Price Increased - Volume Decreased', 'NAME': ' ',
                                   'PREV_CL_PR': '', 'CLOSE_PRICE': '', 'Change': '', 'Change(%)': '',
                                   'Volume Change': '', 'Volume Change(%)': '', 'NET_TRDQTY': '', 'PREV_VOL': ''},
                                  index=[0])

        narration3 = pd.DataFrame({'SYMBOL': 'Price Decreased - Volume Increased', 'NAME': ' ',
                                   'PREV_CL_PR': '', 'CLOSE_PRICE': '', 'Change': '', 'Change(%)': '',
                                   'Volume Change': '', 'Volume Change(%)': '', 'NET_TRDQTY': '', 'PREV_VOL': ''},
                                  index=[0])

        narration4 = pd.DataFrame({'SYMBOL': 'Price Decreased - Volume Decreased', 'NAME': ' ',
                                   'PREV_CL_PR': '', 'CLOSE_PRICE': '', 'Change': '', 'Change(%)': '',
                                   'Volume Change': '', 'Volume Change(%)': '', 'NET_TRDQTY': '', 'PREV_VOL': ''},
                                  index=[0])

        # Append Data Frame when Price Increased and Volume Increased
        narration1 = narration1.append(cust_header)
        price_incr_vol_incr = narration1.append(price_incr_vol_incr, sort=False)

        # Append Data Frame when Price Increased and Volume Increased
        narration2 = narration2.append(cust_header)
        price_incr_vol_decr = narration2.append(price_incr_vol_decr, sort=False)
        price_incr_vol_incr = price_incr_vol_incr.append(price_incr_vol_decr, sort=False)

        # Append Data Frame when Price Increased and Volume Increased
        narration3 = narration3.append(cust_header)
        price_decr_vol_incr = narration3.append(price_decr_vol_incr, sort=False)
        price_incr_vol_incr = price_incr_vol_incr.append(price_decr_vol_incr, sort=False)

        # Append Data Frame when Price Increased and Volume Increased
        narration4 = narration4.append(cust_header)
        price_decr_vol_decr = narration4.append(price_decr_vol_decr, sort=False)
        price_incr_vol_incr = price_incr_vol_incr.append(price_decr_vol_decr, sort=False)

        price_incr_vol_incr = price_incr_vol_incr[['SYMBOL', 'NAME', 'PREV_CL_PR', 'CLOSE_PRICE', 'Change', 'Change(%)',
                                                   'PREV_VOL', 'NET_TRDQTY', 'Volume Change', 'Volume Change(%)']]\
            .copy(deep=True)

        # Format the Excel sheet
        writer = pd.ExcelWriter(report_name, engine='openpyxl')

        # Convert the dataframe to an XlsxWriter Excel object.
        price_incr_vol_incr.to_excel(writer, sheet_name=dp.sheet_name_price_volume, header=None, index=False)

        # Get the openpyxl workbook and worksheet objects.
        worksheet = writer.sheets[dp.sheet_name_price_volume]

        max_rows = worksheet.max_row
        st_row = 1
        st_col = 1
        for row_cells in worksheet.iter_rows(min_row=1, max_row=max_rows):
            for cell in row_cells:
                worksheet.cell(st_row, st_col).border = srgh.thin_border
                if 'Price' in str(cell.value) and 'Volume' in str(cell.value):
                    worksheet.merge_cells(start_row=st_row, start_column=st_col, end_row=st_row, end_column=st_col + 9)
                    worksheet.cell(st_row, st_col).fill = PatternFill(start_color='D3D3D3', fill_type="solid")
                    worksheet.cell(st_row, st_col).font = srgh.font_header
                    worksheet.cell(st_row, st_col).alignment = srgh.align_header
                    break
                elif str(cell.value) in srgh.price_volume_header:
                    worksheet.cell(st_row, st_col).fill = PatternFill(start_color='D3D3D3', fill_type="solid")
                    worksheet.cell(st_row, st_col).font = srgh.font_header
                    worksheet.cell(st_row, st_col).alignment = srgh.align_header
                    st_col = st_col + 1

                else:
                    worksheet.cell(st_row, st_col).font = srgh.font_body
                    if st_row % 2 == 1:
                        cell.fill = PatternFill(start_color="f7ec8f", fill_type="solid")
                    else:
                        cell.fill = PatternFill(start_color="edead7", fill_type="solid")

                    if type(cell.value) == str:
                        worksheet.cell(st_row, st_col).alignment = srgh.align_body_str
                    else:
                        worksheet.cell(st_row, st_col).alignment = srgh.align_body_num
                    st_col = st_col + 1

            st_col = 1
            st_row = st_row + 1
            worksheet.row_dimensions[st_row].height = 20  # In pixels

            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 25
            worksheet.column_dimensions['C'].width = 15
            worksheet.column_dimensions['D'].width = 15
            worksheet.column_dimensions['E'].width = 15
            worksheet.column_dimensions['F'].width = 15
            worksheet.column_dimensions['G'].width = 15
            worksheet.column_dimensions['H'].width = 15
            worksheet.column_dimensions['I'].width = 15
            worksheet.column_dimensions['J'].width = 18
        writer.save()

    # Generate the Price Volume Report
    def generate_new_52_week_high_low_report(self, current_date, report_name):

        master_data = pd.read_excel(self.input_file_name, self.data_sheet_name)
        master_data = master_data[(master_data['TRADE_DATE'] == current_date)]

        master_data = master_data[['SYMBOL', 'NAME', 'HI_52_WK', 'LO_52_WK', 'PREV_CL_PR', 'OPEN_PRICE',
                                   'HIGH_PRICE', 'LOW_PRICE', 'CLOSE_PRICE']]

        new_52_week_high = master_data[(master_data['HI_52_WK'] == master_data['HIGH_PRICE'])]

        new_52_week_low = master_data[(master_data['LO_52_WK'] == master_data['LOW_PRICE'])]

        close_52_week_high = master_data[master_data['CLOSE_PRICE'] >
                                         (master_data['HI_52_WK'] - ((master_data['HI_52_WK'] * 10) / 100))]

        close_52_week_low = master_data[master_data['CLOSE_PRICE'] <
                                        (master_data['LO_52_WK'] + ((master_data['LO_52_WK'] * 10) / 100))]

        narration1 = pd.DataFrame({'SYMBOL': 'New 52 Week High', 'NAME': ' ',
                                   'HI_52_WK': '', 'LO_52_WK': '', 'PREV_CL_PR': '', 'OPEN_PRICE': '',
                                   'HIGH_PRICE': '', 'LOW_PRICE': '', 'CLOSE_PRICE': ''},
                                  index=[0])

        cust_header = pd.DataFrame({'SYMBOL': 'SYMBOL', 'NAME': 'Name', 'HI_52_WK': '52 Week High',
                                    'LO_52_WK': '52 Week Low', 'PREV_CL_PR': 'Previous Close Price',
                                    'OPEN_PRICE': 'Open Price', 'HIGH_PRICE': 'High Price', 'LOW_PRICE': 'Low Price',
                                    'CLOSE_PRICE': 'Close Price'}, index=[0])

        narration2 = pd.DataFrame({'SYMBOL': 'New 52 Week Low', 'NAME': ' ',
                                   'HI_52_WK': '', 'LO_52_WK': '', 'PREV_CL_PR': '', 'OPEN_PRICE': '',
                                   'HIGH_PRICE': '', 'LOW_PRICE': '', 'CLOSE_PRICE': ''},
                                  index=[0])

        narration3 = pd.DataFrame({'SYMBOL': 'Near 52 Week High (10%)', 'NAME': ' ',
                                   'HI_52_WK': '', 'LO_52_WK': '', 'PREV_CL_PR': '', 'OPEN_PRICE': '',
                                   'HIGH_PRICE': '', 'LOW_PRICE': '', 'CLOSE_PRICE': ''},
                                  index=[0])

        narration4 = pd.DataFrame({'SYMBOL': 'Near 52 Week Low (10%)', 'NAME': ' ',
                                   'HI_52_WK': '', 'LO_52_WK': '', 'PREV_CL_PR': '', 'OPEN_PRICE': '',
                                   'HIGH_PRICE': '', 'LOW_PRICE': '', 'CLOSE_PRICE': ''},
                                  index=[0])

        # Append Data Frame when Price Increased and Volume Increased
        narration1 = narration1.append(cust_header)
        new_52_week_high = narration1.append(new_52_week_high, sort=False)

        # Append Data Frame when Price Increased and Volume Increased
        narration2 = narration2.append(cust_header)
        new_52_week_low = narration2.append(new_52_week_low, sort=False)
        new_52_week_high = new_52_week_high.append(new_52_week_low, sort=False)

        # Append Data Frame when Price Increased and Volume Increased
        narration3 = narration3.append(cust_header)
        close_52_week_high = narration3.append(close_52_week_high, sort=False)
        new_52_week_high = new_52_week_high.append(close_52_week_high, sort=False)

        # Append Data Frame when Price Increased and Volume Increased
        narration4 = narration4.append(cust_header)
        close_52_week_low = narration4.append(close_52_week_low, sort=False)
        new_52_week_high = new_52_week_high.append(close_52_week_low, sort=False)

        book = load_workbook(report_name)
        writer = pd.ExcelWriter(report_name, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        new_52_week_high.to_excel(writer, "52-week-high-low", header=None, index=False)

        # Get the openpyxl workbook and worksheet objects.
        worksheet = writer.sheets['52-week-high-low']

        max_rows = worksheet.max_row
        st_row = 1
        st_col = 1
        for row_cells in worksheet.iter_rows(min_row=1, max_row=max_rows):
            for cell in row_cells:
                worksheet.cell(st_row, st_col).border = srgh.thin_border
                if 'New 52 Week' in str(cell.value) or 'Near 52 Week' in str(cell.value):
                    worksheet.merge_cells(start_row=st_row, start_column=st_col, end_row=st_row, end_column=st_col + 8)
                    worksheet.cell(st_row, st_col).fill = PatternFill(start_color='D3D3D3', fill_type="solid")
                    worksheet.cell(st_row, st_col).font = srgh.font_header
                    worksheet.cell(st_row, st_col).alignment = srgh.align_header
                    break
                elif str(cell.value) in srgh.week_high_low_haeder:
                    worksheet.cell(st_row, st_col).fill = PatternFill(start_color='D3D3D3', fill_type="solid")
                    worksheet.cell(st_row, st_col).font = srgh.font_header
                    worksheet.cell(st_row, st_col).alignment = srgh.align_header
                    st_col = st_col + 1

                else:
                    worksheet.cell(st_row, st_col).font = srgh.font_body
                    if st_row % 2 == 1:
                        cell.fill = PatternFill(start_color="f7ec8f", fill_type="solid")
                    else:
                        cell.fill = PatternFill(start_color="edead7", fill_type="solid")

                    if type(cell.value) == str:
                        worksheet.cell(st_row, st_col).alignment = srgh.align_body_str
                    else:
                        worksheet.cell(st_row, st_col).alignment = srgh.align_body_num
                    st_col = st_col + 1

            st_col = 1
            st_row = st_row + 1
            worksheet.row_dimensions[st_row].height = 20  # In pixels

            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 25
            worksheet.column_dimensions['C'].width = 15
            worksheet.column_dimensions['D'].width = 15
            worksheet.column_dimensions['E'].width = 18
            worksheet.column_dimensions['F'].width = 15
            worksheet.column_dimensions['G'].width = 15
            worksheet.column_dimensions['H'].width = 15
            worksheet.column_dimensions['I'].width = 15

        writer.save()

    # Generate the Volatile Stock of the Day Report
    def generate_volatile_stock_day(self, current_date, report_name):

        master_data = pd.read_excel(self.input_file_name, self.data_sheet_name)
        master_data = master_data[(master_data['TRADE_DATE'] == current_date)]

        master_data = master_data[['SYMBOL', 'NAME', 'PREV_CL_PR', 'OPEN_PRICE',
                                   'HIGH_PRICE', 'LOW_PRICE', 'CLOSE_PRICE']]

        high_low_diff = (master_data[master_data.columns[4]] - master_data[master_data.columns[5]])
        high_low_diff_per = ((high_low_diff * 100) / master_data[master_data.columns[5]])

        master_data.insert(5, "Volatility", high_low_diff)
        master_data.insert(6, "Volatility(%)", high_low_diff_per)

        master_data.rename(columns=srgh.volatility_header_updated, inplace=True)

        master_data = master_data[['SYMBOL', 'Name', 'Previous Close Price', 'Open Price',
                                   'High Price', 'Low Price', 'Close Price', 'Volatility', 'Volatility(%)']]

        master_data.sort_values(["Volatility(%)"], axis=0, ascending=False, inplace=True)

        book = load_workbook(report_name)
        writer = pd.ExcelWriter(report_name, engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        master_data.to_excel(writer, "Volatility", index=False)

        # Get the openpyxl workbook and worksheet objects.
        worksheet = writer.sheets['Volatility']

        max_rows = worksheet.max_row
        st_row = 1
        st_col = 1
        for row_cells in worksheet.iter_rows(min_row=1, max_row=max_rows):
            for cell in row_cells:
                worksheet.cell(st_row, st_col).border = srgh.thin_border
                if str(cell.value) in srgh.volatility_header:
                    worksheet.cell(st_row, st_col).fill = PatternFill(start_color='D3D3D3', fill_type="solid")
                    worksheet.cell(st_row, st_col).font = srgh.font_header
                    worksheet.cell(st_row, st_col).alignment = srgh.align_header
                    st_col = st_col + 1

                else:
                    worksheet.cell(st_row, st_col).font = srgh.font_body
                    if st_row % 2 == 1:
                        cell.fill = PatternFill(start_color="f7ec8f", fill_type="solid")
                    else:
                        cell.fill = PatternFill(start_color="edead7", fill_type="solid")

                    if type(cell.value) == str:
                        worksheet.cell(st_row, st_col).alignment = srgh.align_body_str
                    else:
                        worksheet.cell(st_row, st_col).alignment = srgh.align_body_num
                    st_col = st_col + 1

            st_col = 1
            st_row = st_row + 1
            worksheet.row_dimensions[st_row].height = 20  # In pixels

            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 25
            worksheet.column_dimensions['C'].width = 18
            worksheet.column_dimensions['D'].width = 15
            worksheet.column_dimensions['E'].width = 18
            worksheet.column_dimensions['F'].width = 15
            worksheet.column_dimensions['G'].width = 15
            worksheet.column_dimensions['H'].width = 15
            worksheet.column_dimensions['I'].width = 15

            c = worksheet['A2']
            worksheet.freeze_panes = c

        writer.save()

