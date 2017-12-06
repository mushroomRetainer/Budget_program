
from datetime import datetime, timedelta

class Utility:        
    def find_bottom_row(worksheet):
        '''finds the row closest to the bottom of the worksheet that has a non-empty cell'''
        date_values = worksheet.col_values(1)
        bottom_row = len(date_values)
        
        for i in range(bottom_row, 0, -1):
            if date_values[i-1] != None and date_values[i-1] != '': # need to convert from 1- to 0-index
                return i
        return 1
#        current_row = worksheet.row_count
#        while current_row>0:
#            empty_row = True
#            row_values = worksheet.row_values(current_row)
#            for value in row_values:
#                empty_row = empty_row & (value == '')
#            if empty_row:
#                current_row-=1
#            else:
#                break
#        return current_row
        
    def get_row_from_date(worksheet, date, params):
        '''return the row of the first instance of a date that is equal to the given date or the closest date that is before the given date'''
        # in case a date was passed instead of a datetime:
        date = datetime(date.year, date.month, date.day)
        row_counter = 3
        num_rows = Utility.find_bottom_row(worksheet)
        
        date_values = worksheet.col_values(1)
        
        # find the appropriate row to insert the line:
        while row_counter <= num_rows:
#            current_date_str = worksheet.cell(row_counter, 1).value
            current_date_str = date_values[row_counter-1]
            if date <= datetime.strptime(current_date_str, params.date_format):
                break
            row_counter += 1
        return row_counter
        
    def add_budget_line_item(worksheet, date, check_number, bank_description, app_discription, amount, category, params):
        '''adds a budget item to the given worksheet. Determines where to add it based on the date. 
        Assumes there is an existing line of data. Adds it AFTER items with the same date'''
        # just in case date is not a datetime:
        date = datetime(year=date.year,month=date.month,day=date.day)
        
        row_counter = 3
        num_rows = worksheet.row_count
        
        date_values = worksheet.col_values(1)
        
        # find the appropriate row to insert the line:
        while row_counter <= num_rows:
#            current_date_str = worksheet.cell(row_counter, 1).value
            current_date_str = date_values[row_counter-1]
            if current_date_str == '':
                break
            if date < datetime.strptime(current_date_str, params.date_format):
                break
            row_counter += 1
        # copy values from the above row
        values = []
        values.append(date.strftime(params.date_format))
        values.append(check_number)
        values.append(bank_description)
        values.append(app_discription)
        values.append(amount)
        values.append(category)
        worksheet.insert_row(values,index=row_counter)
    
    def get_column_number(worksheet, category):
        '''returns the column number for a given category in a given worksheet'''
        row_values = worksheet.row_values(2)
        if category in row_values:
            return row_values.index(category) + 1 #convert from zero- to one-index
        else:
            return None
        
    def get_current_month_ws(workbook, date, params):
        
        worksheet_names = Utility.get_all_ws_names(workbook)
        
        current_ws_name = params.Month_names[date.month] + ' ' + str(date.year)
        if current_ws_name in worksheet_names:
            return workbook.worksheet(current_ws_name)
        else:
            return None
        
    def get_all_ws_names(workbook):
        all_worksheets = workbook.worksheets()
        worksheet_names = []
        for worksheet in all_worksheets:
            worksheet_names.append(worksheet.title)
        return worksheet_names