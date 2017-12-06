
# python imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import itertools
from collections import Counter # needed for complex merging to preserve duplicates in individual lists, but not in both
import os

# local imports
from params import Params
from input_methods import Import_methods
from utillity import Utility
from supporting_classes import App_entry, Bank_entry, Budget_parameters, Budget_balancer_entry



################################
########### Methods ############
################################

def main(params, fake_date=None):
    '''driving function''' 
    # done: move the json file into a different folder on the pi and on my other computers so it isn't put on the github (might need to delete and recreate the github setup)
    # done: create an offline budget sheet to play with and add an optional parameter to use that sheet for testing new features
    # TODO: write a function that reads in rows in batch and returns a list of lists of cells with row/col indexes 
    # done: alphabetize the final output of the categories on the main page (the appear in a random order each time)
    # TODO: fix the 'previous month net' to actually be the previous month instead of the current month
    # TODO: add support for combined expense logging
    # TODO: add try/except blocks that will output an error message
    # TODO: make a function that uses the budget balancer requests to modify parameters and retroactively change allotments
    # TODO: delete old app entries when the timestamp is more than two weeks behind the latest bank entry
    # TODO: add text message reminders/updates (number of unmatched expenses, pay rent reminder if category is positive, assessment of how we are doing on each category at end of month, etc.)
    # continued: send a reminder/update text at the beginning of the month (or week) with:
    #   how we are doing on categories, 
    #   if we need to do some money transfers, 
    #   remind us to input the latest bank record (if it has been too long)
    #   and how many outstanding items need to be categorized
    # TODO: add a log page that outputs what was done at each update (number of items matched, new month/week created, etc.)
    # rather than one long list. This will make switching out slow reading and writing faster 
    # TODO: make a separate 'trouble shooting scripts file'; add the existing full propagation method; make a new method that takes a csv and does a hard check for duplicates as well as too many of each auto-generated entry as well as dates out of order
    # TODO: read the month of entries in as a batch when comparing for duplicates. Perhaps just have a class that stores all these values?
    # TODO: lots of the output needs to be sorted by date or alphbetically: unresolved input page, monthly/weekly refills
    # TODO: fix the part of the addition propagation that seems to add a second 'copy of end of last month' row. It happens particularily with the full propagation on November, but not October for some reason
    # TODO: paste the unresolved entires in as a batch rather than one at a time
    
    print('Getting Workbook')
    workbook = get_workbook()
    
    # init, every time
    
    if fake_date is None:
        current_datetime = datetime.now()
    else:
        current_datetime = fake_date
        
    print('Determining Last Sync Time')
    last_sync = Import_methods.read_last_sync(workbook, params)
    print('Reading Raw Input')
    app_input_raw = Import_methods.read_app_input(workbook, params)
    bank_data_raw = Import_methods.read_bank_data(workbook, params)
    bank_unresolved_raw = Import_methods.read_bank_unresolved(workbook, params)
    budget_parameters_income_raw, budget_parameters_expenses_raw = Import_methods.read_budget_parameters(workbook, params)
    budget_balancer_input_raw = Import_methods.read_budget_balancer_input(workbook, params)
    
    # these variables are lists of objects from the classes
    print('Cleaning Up Raw Input')
    app_input, bank_data, bank_unresolved, budget_parameters, budget_balancer_input = clean_input_up(app_input_raw, 
                                                                                                     bank_data_raw, 
                                                                                                     bank_unresolved_raw, 
                                                                                                     budget_parameters_income_raw, 
                                                                                                     budget_parameters_expenses_raw, 
                                                                                                     budget_balancer_input_raw)
    print('Removing Bank Data that is too old to use')
    bank_data = delete_old_bank_data(bank_data, params) 
    print('Removing Duplicates from Raw Bank Data that are already known')
    bank_data = remove_duplicate_bank_entries(workbook, bank_data, params)
    print('Merging Raw Bank data with Unresolved Bank Data')
    bank_data = merge_bank_data_and_unresolved(bank_data, bank_unresolved)
    
    is_new_week, is_new_month = is_new_week_or_month(current_datetime, last_sync)
    
    print('Checking for Updates')
    if not (has_updates(app_input, bank_data, budget_balancer_input) or is_new_week or is_new_month):
        print('There are no updates, so the program will now update the sync time and end')
        update_synctime(workbook.worksheet(params.feedpage_ws), current_datetime, params)
        print('Complete')
        return #if there are not updates
    
    earliest_modified_date = current_datetime
    
    if is_new_month:
        # monthly, Call on the first day of the month if the sync value is a previous day
        # for each worksheet, delete_blank_rows() (this helps avoid excessive work when reading in values)
        print('Creating New Month')
        first_modified_date = create_new_month(workbook, current_datetime, last_sync, budget_parameters, params)
        earliest_modified_date = min(earliest_modified_date, first_modified_date)
    
    if is_new_week:
        # weekly. Call on the first day of the week if the sync value is a previous day
        print('Doing Weekly Refills')
        first_modified_date = start_new_week(workbook, current_datetime, last_sync, budget_parameters, params)
        earliest_modified_date = min(earliest_modified_date, first_modified_date)
    
    # every time there are updates:
    print('Assigning Categories')
    
    assign_categories(app_input, bank_data)
    first_modified_date, unmatched_app_entries, total_unresolved = move_and_delete_matches(workbook, app_input, bank_data, current_datetime, params)
    earliest_modified_date = min(earliest_modified_date, first_modified_date)
    print('Earliest Modified Date:',earliest_modified_date)
    
    print('Balancing Budget')
    balance_budget(workbook, budget_balancer_input, current_datetime, params)
    print('Propagating Addition')
    propagate_addition(workbook, earliest_modified_date, params)
    print('Finding Net Changes')
    most_recent_net = recalculate_monthly_net(workbook, earliest_modified_date)
    projected_net = budget_parameters.get_projected_net()
    print('Updating Feedback Page')
    output_dictionary = get_update_page_info(workbook, budget_parameters, unmatched_app_entries, current_datetime)
    update_feedback_page(workbook, current_datetime, total_unresolved, output_dictionary, most_recent_net, projected_net, params)
    
    
#    if is_new_week:
#        possily send a text reminder here too
        
    print('Complete')

def full_propagation(params):
    print('Getting Workbook')
    workbook = get_workbook()
    print('Doing a full propagation of addition')
    propagate_addition(workbook, datetime(year=2017,month=10,day=1), params)

def get_workbook():
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds']
#    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    creds = ServiceAccountCredentials.from_json_keyfile_name(os.path.join(os.pardir, "client_secret_secure.json"), scope) # get json file from parent directory

    client = gspread.authorize(creds)
    if params.run_online:
        print('Using Online Spreadsheet')
        workbook = client.open("Blackburn Budget")
    else:
        print('Using Offline Spreadsheet')
        workbook = client.open("Test Budget")
    return workbook



def is_new_week_or_month(current_datetime, last_sync):
    '''checks to see if a new week or month has started based on the current date/time and the last sync time. Returns two booleans'''
    is_new_week = False
    is_new_month = False
    if (current_datetime - last_sync).days >=7 or current_datetime.weekday() < last_sync.weekday():
        is_new_week = True
    if current_datetime.month > last_sync.month:
        is_new_month = True
    return is_new_week, is_new_month



def clean_input_up(app_input_raw, bank_data_raw, bank_unresolved_raw, budget_parameters_income_raw, budget_parameters_expenses_raw, budget_balancer_input_raw):
    '''cleans the input so there no funny business that happens due to a bad input value. Uses the custom classes to better organize things.'''
    # return app_input, bank_data, budget_parameters, budget_balancer_input
    app_input = []
    bank_data = []
    bank_unresolved = []
    
    budget_balancer_input = []
    
    row_counter = 2 # since there is a header
    for values in app_input_raw:
        if values[0] != '': # check for blank line by looking at the first value
            app_input.append(App_entry(row_counter, values, params))
        row_counter +=1
        
    row_counter = 2 # since there is a header
    for values in bank_data_raw:
        if values[0] != '': # check for blank line by looking at the first value
            bank_data.append(Bank_entry(row_counter, values, params))
        row_counter +=1
        
    row_counter = 2 # since there is a header
    for values in bank_unresolved_raw:
        if values[0] != '': # check for blank line by looking at the first value
            bank_unresolved.append(Bank_entry(row_counter, values, params))
        row_counter +=1
    
    budget_parameters = Budget_parameters(budget_parameters_income_raw, budget_parameters_expenses_raw)
        
    row_counter = 2 # since there is a header
    for values in budget_balancer_input_raw:
        if values[0] != '': # check for blank line by looking at the first value
            budget_balancer_input.append(Budget_balancer_entry(row_counter, values))
        row_counter +=1
        
    return app_input, bank_data, bank_unresolved, budget_parameters, budget_balancer_input

def has_updates(app_input, bank_data, budget_balancer_input):
    '''check to see if there are any new bank entires or app entires or if it is a new week/month. Returns True/False'''
    for app_entry in app_input:
        if not app_entry.included_in_projection:
            return True   
    if len(bank_data) or len(budget_balancer_input) > 0:
        return True
    return False

def create_new_month(workbook, current_datetime, last_sync, budget_parameters, params):
    '''makes a new worksheet titled [month year]. This is only called at the start of a month. Also fills in the category headers'''
    worksheet_names = Utility.get_all_ws_names(workbook)
    first_modified_date = current_datetime
    
    month_counter = last_sync.month
    print("last_sync:",last_sync)
#    print("last_sync.month:",last_sync.month)
    print("current_datetime:",current_datetime)
#    print("current_datetime.month:",current_datetime.month)
    year_counter = last_sync.year
    num_categories = budget_parameters.num_categories
    
    while year_counter <= current_datetime.year and month_counter <= current_datetime.month:
        current_ws_name = params.Month_names[month_counter] + ' ' + str(year_counter)
        if not current_ws_name in worksheet_names:
            worksheet = workbook.add_worksheet(current_ws_name, 3, 6 + num_categories)
            worksheet_names.append(current_ws_name)
            worksheet.update_cell(1,2,"Monthly Net Change:")
            worksheet.update_cell(1,4,"Running Net Change:")
            worksheet.update_cell(2,1,"Date")
            worksheet.update_cell(2,2,"Check Number")
            worksheet.update_cell(2,3,"Bank Description")
            worksheet.update_cell(2,4,"Additional Description")
            worksheet.update_cell(2,5,"Amount")
            worksheet.update_cell(2,6,"Category")
            column_counter = 7
            for category in budget_parameters.all_categories:
                worksheet.update_cell(2,column_counter,category)
                column_counter += 1
                
            # copy values from the previous month over
            previous_month = month_counter - 1
            previous_year = year_counter
            if previous_month < 1:
                previous_month += 12
                previous_year -= 1
            previous_ws_name = params.Month_names[previous_month] + ' ' + str(previous_year)
            
            beginning_of_month = datetime(year_counter, month_counter, 1)
            first_modified_date = min(beginning_of_month,first_modified_date)
            
            if previous_ws_name in worksheet_names:
                previous_ws = workbook.worksheet(previous_ws_name)
                copy_end_of_month_values(previous_ws, worksheet, beginning_of_month, params)
                
            print('Created sheet named:',current_ws_name,'and added initial values.')
            # do monthly refills
            print('Doing Monthly Refills for',current_ws_name)
            monthly_refill(worksheet, beginning_of_month, budget_parameters, params)
        else:
            print('Did not need to create sheet named:',current_ws_name)        
        # increment the month and year counters
        month_counter += 1
        if month_counter > 12:
            month_counter -= 12
            year_counter += 1
            
    return first_modified_date

def copy_end_of_month_values(previous_ws, current_ws, beginning_of_month, params):
    #delete existing row if it exists:
    if current_ws.cell(3, 3) == params.last_month_copy_description:
        current_ws.delete_row(3)
        
    bottom_row = Utility.find_bottom_row(previous_ws)
    end_of_month_values = previous_ws.row_values(bottom_row)
    beginning_of_month_values = [0] * current_ws.col_count # default all values to zero just in case there are new categories
    
    # assign category vlues to the proper category in the next month (may not be the same column)
    for col_counter in range(7, previous_ws.col_count+1):
        new_col_number = Utility.get_column_number(current_ws, previous_ws.cell(2,col_counter).value )
        if new_col_number is not None:
            beginning_of_month_values[new_col_number - 1] =  end_of_month_values[col_counter - 1]
    
    # overwrite the first few lines
    beginning_of_month_values[0] = beginning_of_month.strftime(params.date_format)
    beginning_of_month_values[1] = ''
    beginning_of_month_values[2] = params.last_month_copy_description
    beginning_of_month_values[3] = ''
    beginning_of_month_values[4] = '0'
    beginning_of_month_values[5] = 'auto-generated'
    print('Adding the first row of data to',current_ws.title,'based on',previous_ws.title)
    
    current_ws.insert_row(beginning_of_month_values, index=3)

def propagate_addition(workbook, init_date, params):
    '''starts from the given date and does all the addition from scratch to avoid errors'''
    
    # handle first worksheet
    worksheet = Utility.get_current_month_ws(workbook, init_date, params)
    init_row = Utility.get_row_from_date(worksheet, init_date, params)
    propagate_addition_ws(worksheet, init_row)
    
    # handle subsequent worksheets
    month_counter = init_date.month + 1
    year_counter = init_date.year
    if month_counter > 12:
        month_counter-=12
        year_counter+=1
    date_counter = datetime(year_counter, month_counter, 1) # the first day of next month
    previous_ws = worksheet
    worksheet = Utility.get_current_month_ws(workbook, date_counter, params)
    while worksheet is not None:
        # update beginning of the month values
        copy_end_of_month_values(previous_ws, worksheet, date_counter, params)
        propagate_addition_ws(worksheet, 4) #we don't want to modify the first entry that had the values from last month
        
        # increment to the next first day of the month
        month_counter += 1
        if month_counter > 12:
            month_counter-=12
            year_counter+=1
        date_counter = datetime(year_counter, month_counter, 1)
        
        # update previous and current worksheets
        previous_ws = worksheet
        worksheet = Utility.get_current_month_ws(workbook, date_counter, params)

def propagate_addition_ws(worksheet, init_row):
    num_rows = Utility.find_bottom_row(worksheet)
    num_cols = worksheet.col_count
    if init_row<=3:
        init_row =4 # cannot start before the first line of data
    
#    init_row -= 1
#    cell_list = worksheet.range(init_row, 5, num_rows, num_cols)
#    cell_matrix = []
#    width = num_cols - 4
#    while len(cell_list) > 0:
#        cell_matrix.append(cell_list[0:width])
#        del cell_list[0:width]
#    
#    for row_index in range(1,len(cell_matrix)): # don't use the first row since we don't want to modify the start of the month values
##        print('current row index:',row_index)
#        category = cell_matrix[row_index][1].value
#        amount = float(cell_matrix[row_index][0].value)
#        affected_column = Utility.get_column_number(worksheet, category)
#        row_length = len(cell_matrix[row_index])
#        for col_index in range(2, row_length): # need to neglect the first two columns since they are the amount and category
#            cell_matrix[row_index][col_index].value = cell_matrix[row_index-1][col_index].value
##            print('updated the value of cell',row_index,col_index,':', cell_matrix[row_index][col_index].value)
#            if col_index+1+4 == affected_column: # need to account for the missing columns that were neglected in the index slice and for the 0- to 1- indexing 
#                cell_matrix[row_index][col_index].value = str( float(cell_matrix[row_index][col_index].value) + amount )
#    cell_list = []
#    for row in cell_matrix:
#        cell_list += row
#    worksheet.update_cells(cell_list)
    
    
    previous_values = worksheet.row_values(init_row-1)
    
    for row_counter in range(init_row, num_rows+1):
#        print('row:',row_counter)
        values = worksheet.row_values(row_counter)
#        print('these are the values at row_counter',row_counter,':',values)
        amount = float(values[4])
        category = values[5]
        affected_column = Utility.get_column_number(worksheet, category)
        cells_to_update = worksheet.range(row_counter, 7, row_counter, num_cols)
        column_counter = 7
        for cell in cells_to_update:
            new_value = float(previous_values[column_counter-1]) #account for zero indexing
            if column_counter == affected_column:
                new_value += amount
            cell.value = str(new_value)
            column_counter+=1
        worksheet.update_cells(cells_to_update)
        previous_values = worksheet.row_values(row_counter)  
        
##############################################################################
def clean_up_form_entries():
    '''deletes form entries over a month old (Based on timestamp, not custom date) even if they haven't been used'''
    # make sure not to delete the last unfrozen row
    pass

def start_new_week(workbook, current_datetime, last_sync, budget_parameters, params):
    first_modified_date = current_datetime
    weekday_of_last_sync = last_sync.weekday()
    current_first_weekday = datetime(last_sync.year,last_sync.month, last_sync.day) + timedelta(days=7-weekday_of_last_sync)
    week_increment = timedelta(days=7)
    while current_first_weekday <= current_datetime:
        first_modified_date = min(first_modified_date, current_first_weekday)
        worksheet = Utility.get_current_month_ws(workbook, current_first_weekday, params)
        weekly_refill(worksheet, current_first_weekday, budget_parameters, params)
        current_first_weekday += week_increment
        
    return first_modified_date

def weekly_refill(worksheet, first_date_of_week, budget_parameters, params):
    '''puts weekly refills into each appropriate budget category'''
    month=first_date_of_week.month +1
    year=first_date_of_week.year
    if month > 12:
        month -= 12
        year += 1
    
    first_day_next_month = datetime(year,month, day=1)
    days_left_in_month = (first_day_next_month-first_date_of_week).days
    if days_left_in_month < 7:
        partial_weekly_refill(worksheet, first_date_of_week, days_left_in_month, budget_parameters, params)
    else:
        # income is negative (since it hasn't happened yet)
        for category in budget_parameters.data_layered_dictionary['Income']['weekly'].keys():
            amount = budget_parameters.data_layered_dictionary['Income']['weekly'][category]
            Utility.add_budget_line_item(worksheet, first_date_of_week, '', 'Weekly Refill for '+category, '', -1*amount, category, params)
        # expenses are positive
        for category in budget_parameters.data_layered_dictionary['Expenses']['weekly'].keys():
            amount = budget_parameters.data_layered_dictionary['Expenses']['weekly'][category]
            Utility.add_budget_line_item(worksheet, first_date_of_week, '', 'Weekly Refill for '+category, '', amount, category, params)

def partial_weekly_refill(worksheet, date, num_days, budget_parameters, params):
    partial = num_days/7.0
    # income is negative (since it hasn't happened yet)
    for category in budget_parameters.data_layered_dictionary['Income']['weekly'].keys():
        amount = budget_parameters.data_layered_dictionary['Income']['weekly'][category]
        amount = round(amount*partial,2)
        Utility.add_budget_line_item(worksheet, date, '', 'Partial Weekly Refill for '+category, '', -1*amount, category, params)
    # expenses are positive
    for category in budget_parameters.data_layered_dictionary['Expenses']['weekly'].keys():
        amount = budget_parameters.data_layered_dictionary['Expenses']['weekly'][category]
        amount = round(amount*partial,2)
        Utility.add_budget_line_item(worksheet, date, '', 'Partial Weekly Refill for '+category, '', amount, category, params)

def monthly_refill(worksheet, first_date_of_month, budget_parameters, params):
    '''puts monthly refills into each appropriate budget category'''
    # income is negative (since it hasn't happened yet)
    for category in budget_parameters.data_layered_dictionary['Income']['monthly'].keys():
        amount = budget_parameters.data_layered_dictionary['Income']['monthly'][category]
        Utility.add_budget_line_item(worksheet, first_date_of_month, '', 'Monthly Refill for '+category, '', -1*amount, category)
    # expenses are positive
    for category in budget_parameters.data_layered_dictionary['Expenses']['monthly'].keys():
        amount = budget_parameters.data_layered_dictionary['Expenses']['monthly'][category]
        Utility.add_budget_line_item(worksheet, first_date_of_month, '', 'Monthly Refill for '+category, '', amount, category)
    # do partial weekly refills
    days_in_week = 7 - first_date_of_month.weekday()
    partial_weekly_refill(worksheet, first_date_of_month, days_in_week, budget_parameters, params)
    



def move_and_delete_matches(workbook, app_input, bank_data, current_datetime, params):
    '''takes the matched bank entries and app entries, then inserts bank entries into the appropriate worksheets 
    and deletes the old information from the input sheets and moves outstanding bank entries into their own worksheet
    Also marks app entries that are left as 'Included' '''
    
    worksheet_bank_unresolved = workbook.worksheet(params.unresolved_items_ws)
    first_modified_date = current_datetime
    total_unresolved = 0
    
    # delete all matched app_entries
    rows_to_delete = []
    unmatched_app_entries = []
    
    worksheet_app_entries = workbook.worksheet(params.expense_input_ws)
    for app_entry in app_input:
        if app_entry.is_matched:
            rows_to_delete.append(app_entry.row)
        else:
            unmatched_app_entries.append(app_entry)
            worksheet_app_entries.update_cell(app_entry.row,13,'Included')
    rows_to_delete = sorted(rows_to_delete,reverse=True) # by deleting them in reverse order, we won't have any problems with the rows shifting
    
    # can't delete all unfrozen rows in worksheet, so add a fresh row if needed
    if len(rows_to_delete) == worksheet_app_entries.row_count -1:
        worksheet_app_entries.add_rows(1)
        
    for row in rows_to_delete:
        worksheet_app_entries.delete_row(row)
    
        
    # delete all bank entries
    worksheet_bank_entries = workbook.worksheet(params.raw_bank_data_ws)
    num_rows = worksheet_bank_entries.row_count
    for row in range(2,num_rows+1):
        worksheet_bank_entries.delete_row(2)
    worksheet_bank_entries.add_rows(1) # just to leave one blank row at the bottom so it's easier to past in data
    
    #delete all unresolved items
    num_rows = worksheet_bank_unresolved.row_count
    for row in range(2,num_rows+1):
        worksheet_bank_unresolved.delete_row(2)
        
    #delete all budget app items
    worksheet_budget_balancer = workbook.worksheet(params.budget_balancer_ws)
    num_rows = worksheet_budget_balancer.row_count
    worksheet_budget_balancer.add_rows(1) # add an extra blank row so there is still one left
    for row in range(2,num_rows+1):
        worksheet_budget_balancer.delete_row(2)
      
    # input all bank_entries that are unmatched and matched into appropriate sheets    
    
    bank_data.sort()

    for bank_entry in bank_data:
        if not bank_entry.is_matched:
            values = bank_entry.original_values
            worksheet_bank_unresolved.append_row(values)
            total_unresolved += 1
        else:
            worksheet = Utility.get_current_month_ws(workbook, bank_entry.date, params)
            Utility.add_budget_line_item(worksheet, 
                                         bank_entry.date, 
                                         bank_entry.check_number, 
                                         bank_entry.description, 
                                         bank_entry.app_description, 
                                         bank_entry.amount, 
                                         bank_entry.category)
            # convert date to datetime
            bank_date = datetime(year=bank_entry.date.year, month=bank_entry.date.month, day=bank_entry.date.day)
            first_modified_date = min(first_modified_date, bank_date)
    
    return first_modified_date, unmatched_app_entries, total_unresolved

def remove_duplicate_bank_entries(workbook, bank_data, params):
    '''before bank entries are assigned, we need to make sure they haven't already been inputted and 
    categorized (or put in the outstanding worksheet)'''
    # TODO: this will need to be modified to account for multi-category expenses (since the amount will not be the same)
    # My current idea to do this is to put a pre-fix on the bank description with a flag and the total amount, so we can still check it here
    
    bank_entries_to_keep = []
    entires_matched = [] # this will help handle dulplicates by tracking wich ones have been used to make a match
    for bank_entry in bank_data:
        duplicate = False
        worksheet = Utility.get_current_month_ws(workbook, bank_entry.date, params)
        if worksheet is None:
            bank_entries_to_keep.append(bank_entry)
            continue 
        row = Utility.get_row_from_date(worksheet, bank_entry.date, params)
        num_rows = Utility.find_bottom_row(worksheet)
        if row <= num_rows:
            compare_date = datetime.strptime(worksheet.cell(row,1).value, params.date_format).date()
            while bank_entry.date == compare_date:
                if bank_entry.description == worksheet.cell(row,3).value and bank_entry.amount == float(worksheet.cell(row,5).value): # need to compare amounts as floats, otherwise having/not having a decimal can throw things off if you compare strings
                    ID = [worksheet.title,row]
                    if ID not in entires_matched:
                        duplicate = True
                        entires_matched.append(ID)
                        break
                row +=1
                if row > num_rows:
                    break 
                compare_date = datetime.strptime(worksheet.cell(row,1).value, params.date_format).date()
        if not duplicate:
            bank_entries_to_keep.append(bank_entry)
    return bank_entries_to_keep

def merge_bank_data_and_unresolved(bank_data, bank_unresolved):
    # https://stackoverflow.com/questions/36349881/union-of-two-lists-in-python
    # both objects need a __hash__ method, and you can just return the built in hash() result
    a, b = bank_data, bank_unresolved
    na, nb = Counter(a), Counter(b)
    return list(Counter({k: max((na[k], nb[k])) for k in set(a + b)}).elements())

def delete_old_bank_data(bank_data, params):
    '''removes all bank data that is older than the earliest_tracking_date'''
    bank_data_to_keep = []
    for bank_entry in bank_data:
        if bank_entry.date >= params.earliest_tracking_date:
            bank_data_to_keep.append(bank_entry)
    return bank_data_to_keep

def assign_categories(app_input, bank_data):
    '''matches app input and automatic expenses with new and outstanding bank data. Doesn't change with the actual sheets.'''
    # attempt exact matches
    for bank_entry in bank_data:
        for app_entry in app_input:
             bank_entry.attempt_match_exact(app_entry)
    # attempt approximate matches
    for bank_entry in bank_data:
        for app_entry in app_input:
            bank_entry.attempt_match_approximate(app_entry)
    # attempt combination exact matches
    for bank_entry in bank_data:
        bank_entry.attempt_match_combined_exact(app_input)
    # attempt approximate exact matches
    for bank_entry in bank_data:
        bank_entry.attempt_match_combined_approximate(app_input)
        
def recalculate_monthly_net(workbook, earliest_modified_date):
    most_recent_net = None
    all_worksheets = Utility.get_all_ws_names(workbook)
    date_counter = earliest_modified_date.replace(day=1,hour=0,minute=0,second=0)
    worksheet_name = params.Month_names[date_counter.month] + ' ' + str(date_counter.year)
    while worksheet_name in all_worksheets:
        worksheet = Utility.get_current_month_ws(workbook, date_counter, params)
        first_row = worksheet.row_values(3)[6:]
        last_row = worksheet.row_values(Utility.find_bottom_row(worksheet))[6:]
        first_row_sum = sum(str_list_to_float(first_row)) 
        last_row_sum = sum(str_list_to_float(last_row))
        worksheet.update_cell(1,3,str(last_row_sum - first_row_sum))
        worksheet.update_cell(1,5,str(last_row_sum))
        most_recent_net = last_row_sum
        # increment
        if date_counter.month<12:
            date_counter = date_counter.replace(month=date_counter.month+1)
        else:
            date_counter = date_counter.replace(year=date_counter.year+1,month=1)
        worksheet_name = params.Month_names[date_counter.month] + ' ' + str(date_counter.year)
    return most_recent_net 

def str_list_to_float(strings):
    '''converts all strings to floats, https://stackoverflow.com/questions/7368789/convert-all-strings-in-a-list-to-int'''
    return list(map(float, strings)) 

def get_update_page_info(workbook, budget_parameters, unmatched_app_entries, current_datetime):
    '''returns a clean dictionary with "category:[amount remaining, next refill date, refill amount]"
    also adds in the unmatched app entries too'''
    worksheet = Utility.get_current_month_ws(workbook, current_datetime, params)
    next_first_weekday = datetime(current_datetime.year,current_datetime.month, current_datetime.day) + timedelta(days=7-current_datetime.weekday())
#    next_first_weekday_str = next_first_weekday.strftime(date_format)
    if current_datetime.month < 12:
        next_first_month = datetime(current_datetime.year,current_datetime.month + 1, 1)
    else:
        next_first_month = datetime(current_datetime.year + 1, 1, 1)
#    next_first_month_str = next_first_month.strftime(date_format)
    category_row = worksheet.row_values(2)[6:]
    last_row = worksheet.row_values(Utility.find_bottom_row(worksheet))[6:]
    last_row_values = str_list_to_float(last_row)
    output_dictionary = {}
    index = 0 
    for category in category_row:
        amount = last_row_values[index]
        if category in budget_parameters.data_layered_dictionary['Expenses']['weekly']:
            refill_amount = budget_parameters.data_layered_dictionary['Expenses']['weekly'][category]
            output_dictionary[category] = [amount, next_first_weekday, refill_amount]
        elif category in budget_parameters.data_layered_dictionary['Expenses']['monthly']:
            refill_amount = budget_parameters.data_layered_dictionary['Expenses']['monthly'][category]
            output_dictionary[category] = [amount, next_first_month, refill_amount]
        index += 1
        
    # adjust values based on unmatched app entries
    for app_entry in unmatched_app_entries:
        if app_entry.category in output_dictionary:
            output_dictionary[app_entry.category][0] += app_entry.amount
    
    return output_dictionary

def update_feedback_page(workbook, current_datetime, total_unresolved, output_dictionary, most_recent_net, projected_net, params):
    '''clears the feedback page and updates it with the new budget category values in one step. Also recrods the "sync" time and date '''
    worksheet = workbook.worksheet(params.feedpage_ws)
    
    num_rows = worksheet.row_count
    num_categories = len(output_dictionary)
    if num_rows > num_categories + 5:
        for i in range(num_rows - (num_categories + 5)):
            worksheet.delete_row(6)
    elif num_rows < num_categories + 5:
        worksheet.add_rows(num_categories + 5 - num_rows)
    num_rows = worksheet.row_count
    
    # assign values to cell list
    cell_list = worksheet.range(6, 1, num_rows, 4)
    cell_counter = 0
    categories = list(output_dictionary.keys())
    categories.sort()
    for category in categories:
        values =  output_dictionary[category]
        
        cell_list[cell_counter].value = category
        cell_counter += 1
        
        cell_list[cell_counter].value = str(values[0])
        cell_counter += 1
        
        cell_list[cell_counter].value = values[1].strftime(params.date_format)
        cell_counter += 1
        
        cell_list[cell_counter].value = str(values[2])
        cell_counter += 1
    
    update_synctime(worksheet, current_datetime, params)
    # update net values
    worksheet.update_cell(2,2,str(most_recent_net))
    
    worksheet.update_cell(3,2,str(projected_net[0]))
    worksheet.update_cell(3,4,str(projected_net[1]))
    # update expense categories
    worksheet.update_cells(cell_list)  
    
def update_synctime(worksheet, current_datetime, params):
    # update sync time
    print('New Sync Time is:',current_datetime.strftime(params.datetime_format))
    worksheet.update_cell(1,2,current_datetime.strftime(params.datetime_format))

def balance_budget(workbook, budget_balancer_input, current_datetime, params):
    '''manages input from the budget balancer form. Allows you to modify budget parameters permenently, retroacticely, or just one-time. Also allows you to do one-time transfers between expense categories. Also delete the rows of the form entries that are used'''
    worksheet = Utility.get_current_month_ws(workbook, current_datetime, params)
    for item in budget_balancer_input:
        if item.is_transfer:
            Utility.add_budget_line_item(worksheet, current_datetime, '', 'Budget Balancer Decrease', item.notes, item.amount, item.category)
            Utility.add_budget_line_item(worksheet, current_datetime, '', 'Budget Balancer Increase', item.notes, item.amount_2, item.category_2)
        else: # for permanent budget category adjustments
            pass # TODO: need to write this! The permanent budget category change doesn't work

def output_update(message, workbook = None):
    if workbook is None:
        workbook = get_workbook()
    worksheet = workbook.worksheet(params.feedpage_ws)
    worksheet.update_cell(1,3,message)

#############################################################################
#############################################################################

params = Params()

# use this to redo all the addition from 10-1-17 and on
#full_propagation(params)
#print('successfully propagated addition from the beginning')

if params.run_online:
    try:
        output_update('Currently Working')
        main(params)
        output_update('Complete')
    except:
        print('Program crashed :(')
        try:
            output_update('Program crashed')
        except:
            print('Even the feedback crashed')
else:
    output_update('Currently Working')
    main(params)
    output_update('Complete')
#    main(fake_date=datetime(year=2017,month=9,day=5))

