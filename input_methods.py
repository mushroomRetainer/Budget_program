
from datetime import datetime, timedelta

class Import_methods:
    def read_last_sync(workbook, params):
        '''reads in the date and time the shreadsheet was most recently synconized'''
        worksheet = workbook.worksheet(params.feedpage_ws)
        last_sync_string = worksheet.cell(1,2).value
        last_sync = datetime.strptime(last_sync_string, params.datetime_format)
        return last_sync
    
    def read_app_input(workbook, params):
        '''read in all the current app input'''
        worksheet = workbook.worksheet(params.expense_input_ws)
        num_rows = worksheet.row_count
        app_input_raw = []
        for row_counter in range(1, num_rows+1):
            values = worksheet.row_values(row_counter)
            app_input_raw.append(values)
        # delete headers
        app_input_raw.pop(0)
    
        return app_input_raw
    
    def read_bank_data(workbook, params):
        '''read in all the current bank data'''
        worksheet = workbook.worksheet(params.raw_bank_data_ws)
        num_rows = worksheet.row_count
        bank_data_raw = []
        for row_counter in range(1, num_rows+1):
            values = worksheet.row_values(row_counter)
            bank_data_raw.append(values)
        # delete headers
        bank_data_raw.pop(0)
        
        return bank_data_raw
    
    def read_bank_unresolved(workbook, params):
        '''read in all the current bank data'''
        worksheet = workbook.worksheet(params.unresolved_items_ws)
        num_rows = worksheet.row_count
        bank_unresolved_raw = []
        for row_counter in range(1, num_rows+1):
            values = worksheet.row_values(row_counter)
            bank_unresolved_raw.append(values)
        # delete headers
        bank_unresolved_raw.pop(0)
        
        return bank_unresolved_raw
    
    
    def read_budget_parameters(workbook, params):
        '''read in all budget parameters'''
        worksheet = workbook.worksheet(params.parameters_ws)
        num_cols = worksheet.col_count
        budget_parameters_cols = [[],[],[],[]]
        # need to combine the separate columns into only 4 columns
        num_income_categories = 0
        column_combiner = 0
        for col_counter in range(4, num_cols+1):
            values = worksheet.col_values(col_counter)
            
            # delete headers
            values.pop(0)
            
            budget_parameters_cols[column_combiner] += values
            if col_counter == 4:
                num_income_categories = len(budget_parameters_cols[0])
            column_combiner = (column_combiner+1)%4
        
        budget_parameters_income_raw = []
        budget_parameters_expenses_raw = []
        
        num_rows = len(budget_parameters_cols[0])
        for parameter_index in range(num_rows):
            row = [budget_parameters_cols[0][parameter_index],budget_parameters_cols[1][parameter_index],budget_parameters_cols[2][parameter_index],budget_parameters_cols[3][parameter_index]]
            row_is_empty = True
            for element in row:
                row_is_empty = row_is_empty & (element=='')
            if not row_is_empty:
                if parameter_index < num_income_categories:
                    budget_parameters_income_raw.append(row)
                else:
                    budget_parameters_expenses_raw.append(row)
        return budget_parameters_income_raw, budget_parameters_expenses_raw
    
    def read_budget_balancer_input(workbook, params):
        '''read in all budget balancer input'''
        worksheet = workbook.worksheet(params.budget_balancer_ws)
        num_rows = worksheet.row_count
        budget_balancer_input_raw = []
        for row_counter in range(1, num_rows+1):
            values = worksheet.row_values(row_counter)
            budget_balancer_input_raw.append(values)
    
        # delete headers
        budget_balancer_input_raw.pop(0)
        
        return budget_balancer_input_raw