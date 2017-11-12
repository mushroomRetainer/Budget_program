import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import itertools
from collections import Counter # needed for complex merging to preserve duplicates in individual lists, but not in both
import os

################################
########### Example ############
################################

#worksheet = workbook.worksheet("New Input From App")
# 
## Extract and print all of the values
# num_rows = worksheet.row_count
# for row_counter in range(1, num_rows+1):
#    values = worksheet.row_values(row_counter)
#    print(values)



# adjustable parameters
run_online = False

Month_names = ['month0','Jan','Feb','March','April','May','June','July','Aug','Sept','Oct','Nov','Dec']
feedpage_ws = 'Current Budget'
expense_input_ws = 'Input From Expense Log'
balancer_input_ws = 'Input From Budget Balancer'
parameters_ws = 'Budget Parameters'
budget_balancer_ws = 'Input From Budget Balancer'
auto_transactions_ws = 'Automatic transactions'
raw_bank_data_ws = 'raw bank acount data'
unresolved_items_ws = 'unresolved bank items'
datetime_format = '%m/%d/%Y %H:%M:%S'
date_format = '%m/%d/%Y' # hopefully the same as BECU to cut down on formatting
last_month_copy_description = 'Copy from End of Last Month'
date_match_tolerance = timedelta(days=3)
earliest_tracking_date = datetime(year=2017,month=10,day=1).date() # anything before this will be 
# other sheets are created and named dynamically based on month/year



################################
########### Methods ############
################################

#def test_read_matrix(workbook):
#    ## Select a range
#    #cell_list = worksheet.range('A1:C7')
#    cell_list = workbook.worksheet('test sheet').range(2, 2, 6, 6)
#    counter = 101
#    for cell in cell_list:
#        print(cell.value)
#        cell.value = counter
#        counter+=1
#    # Update in batch
#    workbook.worksheet('test sheet').update_cells(cell_list)   

def main(fake_date=None):
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
    print('Getting Workbook')
    workbook = get_workbook()
    
#    print('testing mass read/write')
#    test_read_matrix(workbook)
    
    # init, every time
    
    if fake_date is None:
        current_datetime = datetime.now()
    else:
        current_datetime = fake_date
        
    print('Determining Last Sync Time')
    last_sync = read_last_sync(workbook)
    print('Reading Raw Input')
    app_input_raw = read_app_input(workbook)
    bank_data_raw = read_bank_data(workbook)
    bank_unresolved_raw = read_bank_unresolved(workbook)
    budget_parameters_income_raw, budget_parameters_expenses_raw = read_budget_parameters(workbook)
    budget_balancer_input_raw = read_budget_balancer_input(workbook)
    
    # these variables are lists of objects from the classes
    print('Cleaning Up and Organizing Raw Input')
    app_input, bank_data, bank_unresolved, budget_parameters, budget_balancer_input = clean_input_up(app_input_raw, bank_data_raw, bank_unresolved_raw, budget_parameters_income_raw, budget_parameters_expenses_raw, budget_balancer_input_raw)
    bank_data = delete_old_bank_data(bank_data) 
    bank_data = remove_duplicate_bank_entries(workbook, bank_data)
    bank_data = merge_bank_data_and_unresolved(bank_data, bank_unresolved)
    
    is_new_week, is_new_month = is_new_week_or_month(current_datetime, last_sync)
    
    print('Checking for Updates')
    if not (has_updates(app_input, bank_data, budget_balancer_input) or is_new_week or is_new_month):
        return #if there are not updates
    
    earliest_modified_date = current_datetime
    
    if is_new_month:
        # monthly, Call on the first day of the month if the sync value is a previous day
        # for each worksheet, delete_blank_rows() (this helps avoid excessive work when reading in values)
        print('Creating New Month')
        first_modified_date = create_new_month(workbook, current_datetime, last_sync, budget_parameters)
        earliest_modified_date = min(earliest_modified_date, first_modified_date)
    
    if is_new_week:
        # weekly. Call on the first day of the week if the sync value is a previous day
        print('Doing Weekly Refills')
        first_modified_date = start_new_week(workbook, current_datetime, last_sync, budget_parameters)
        earliest_modified_date = min(earliest_modified_date, first_modified_date)
    
    # every time there are updates:
    print('Assigning Categories')
    
    assign_categories(app_input, bank_data)
    first_modified_date, unmatched_app_entries, total_unresolved = move_and_delete_matches(workbook, app_input, bank_data, current_datetime)
    earliest_modified_date = min(earliest_modified_date, first_modified_date)
    print('Earliest Modified Date:',earliest_modified_date)
    print('Propagating Addition')
    propagate_addition(workbook, earliest_modified_date)
    print('Finding Net Changes')
    most_recent_net = recalculate_monthly_net(workbook, earliest_modified_date)
    projected_net = budget_parameters.get_projected_net()
    print('Updating Feedback Page')
    output_dictionary = get_update_page_info(workbook, budget_parameters, unmatched_app_entries, current_datetime)
    update_feedback_page(workbook, current_datetime, total_unresolved, output_dictionary, most_recent_net, projected_net)
    print('Balancing Budget')
    balance_budget()
    
#    if is_new_week:
#        clean_up_form_entries() # this needs to happen after the category assignments in case you assigned a very retroactive one
#        possily send a text reminder here too
        
    
    print('Complete')

class App_entry:
    ## order of the entries in "values":
    #0 Timestamp	
    #1 Who made the transaction?	
    #2 Category	
    #3 Date (today is assumed if left blank)	
    #4 Income Sub-category	
    #5 Donations & Savings Sub-category	
    #6 Bills Sub-category	
    #7 Daily Living Sub-category	
    #8 Brief Description (if necessary)	
    #9 Check if This is a Refund Applied to an Expense Category
    #10 Enter Amount of Transaction	
    #11 Check if This is Combined With Another Expense On the Same Date	
    #12 Included in projection?
    
    # class attributes:
    # self.row
    # self.date
    # self.who
    # self.category
    # self.description
    # self.is_refund
    # self.amount
    # self.is_combined
    # self.included_in_projection
    
    # self.is_matched
    # self.original_values
    
    def __init__(self, row, values):
        self.row = row # keep in mind this is the spreadsheet row, so it is 1-indexed
        self.original_values = values
        
        if values[3] != '':
            self.date = datetime.strptime(values[3], date_format).date()
        elif values[0] != '':
            self.date = datetime.strptime(values[0], datetime_format).date()
        
        self.who = values[1]
        
        # get the category using one of the four "subcategories"
        if values[2] == 'Income':
            self.category = values[4]
        
        elif values[2] == 'Donations & Savings':
            self.category = values[5]
        
        elif values[2] == 'Bills':
            self.category = values[6]
        
        elif values[2] == 'Daily Living':
            self.category = values[7]
            
        else:
            self.category = '' # not sure why it would get to this poit, but just in case we'll set a empty value
        
        self.description = values[8]
        self.is_refund = (values[9] != '')
        
        # determine positive or negtative sign for the amount (only negative if expense that isn't a refund):
        if values[2] == 'Income' or self.is_refund:
            self.amount = abs(float(values[10]))
        else:
            self.amount =  -1 * abs(float(values[10]))
        
        self.is_combined = (values[11] != '')
        self.included_in_projection = (values[12] != '')
        
        self.is_matched = False
        
    # comparison methods (so a list can be sorted by date)
    def __lt__(self, other):
        return self.date < other.date
    def __le__(self, other):
        return self.date <= other.date
    def __eq__(self, other):
        return self.date == other.date
    def __ne__(self, other):
        return self.date != other.date
    def __gt__(self, other):
        return self.date > other.date
    def __ge__(self, other):
        return self.date >= other.date

class Bank_entry:
    #0 date
    #1 check number
    #2 bank description
    #3 negative
    #4 positive
    
    ## class attributes:
    # self.row
    # self.date
    # self.check_number
    # self.description
    # self.amount
    
    # self.is_matched
    # self.category
    # self.app_description
    # self.original_values
    
    # TODO: implement these:
    # self.is_combined
    # self.categories
    # self.amounts
    
    def __init__(self, row, values):
        self.row = row
        self.original_values = values
        self.date = datetime.strptime(values[0], date_format).date()
        self.check_number = values[1]
        self.description = values[2]
        
        if values[4] != '':
            # income (positive value)
            self.amount = abs(float(values[4]))
        elif values[3] != '':
            # expense (negative value)
            self.amount = -1 * abs(float(values[3]))
        else:
            self.amount=0 # not sure why it would ever get here
        
        self.is_matched = False
        self.category = None
        self.app_description = None
            
    def attempt_match_exact(self, app_entry):
#        print('Exact Match attempt:', self.date, '==', app_entry.date, 'and', self.amount, '==', app_entry.amount)
        if self.is_matched or app_entry.is_matched or app_entry.is_combined:
            return False
        elif self.date == app_entry.date and self.amount == app_entry.amount:
            self.category = app_entry.category
            self.is_matched = True
            self.app_description = app_entry.who + ": " + app_entry.description
            app_entry.is_matched = True
            return True
        else:
            return False
    
    def attempt_match_approximate(self, app_entry):
        if self.is_matched or app_entry.is_matched or app_entry.is_combined:
            return False
        elif abs(self.date - app_entry.date) <= date_match_tolerance and self.amount == app_entry.amount:
            self.category = app_entry.category
            self.is_matched = True
            self.app_description = app_entry.who + ": " + app_entry.description
            app_entry.is_matched = True
            return True
        else:
            return False
    
    def attempt_match_combined_exact(self, app_entry_list):
        if self.is_matched:
            return False
        #get subset of app_entry list that are combined and have the exact date
        app_entry_list_combined = []
        for app_entry in app_entry_list:
            if app_entry.is_combined and self.date == app_entry.date:
                app_entry_list_combined.append(app_entry)
        if len(app_entry_list_combined) == 0:
            return False
        # try all possible combinations and see if one works
        # it is called a powerset when you get all combinations of all lengths
        # https://stackoverflow.com/questions/464864/how-to-get-all-possible-combinations-of-a-list-s-elements
        for L in range(0, len(app_entry_list_combined)+1):
            for subset in itertools.combinations(app_entry_list_combined, L):
                # get sum of subset
                amount_sum = 0
                for app_entry in subset:
                    amount_sum += app_entry.amount
                # if sum is equal, then set matched variables to true and break out of both loops with a return
                if amount_sum == self.amount:
                    self.category = app_entry.category # TODO: This only records the category as the last entry's category! 
                    self.is_matched = True
                    self.app_description = ''
                    for app_entry in subset:
                        self.app_description += app_entry.who + ": " + app_entry.description + '. '
                        app_entry.is_matched = True
                    return True
        return False
    
    def attempt_match_combined_approximate(self, app_entry_list):
        if self.is_matched:
            return False
        #get subset of app_entry list that are combined and have the exact date
        app_entry_list_combined = []
        for app_entry in app_entry_list:
            if app_entry.is_combined and abs(self.date - app_entry.date) <= date_match_tolerance:
                app_entry_list_combined.append(app_entry)
        if len(app_entry_list_combined) == 0:
            return False
        # try all possible combinations and see if one works
        # it is called a powerset when you get all combinations of all lengths
        # https://stackoverflow.com/questions/464864/how-to-get-all-possible-combinations-of-a-list-s-elements
        for L in range(0, len(app_entry_list_combined)+1):
            for subset in itertools.combinations(app_entry_list_combined, L):
                # get sum of subset
                amount_sum = 0
                for app_entry in subset:
                    amount_sum += app_entry.amount
                # if sum is equal, then set matched variables to true and break out of both loops with a return
                if amount_sum == self.amount:
                    self.category = app_entry.category
                    self.is_matched = True
                    self.app_description = ''
                    for app_entry in subset:
                        self.app_description += app_entry.who + ": " + app_entry.description + '. '
                        app_entry.is_matched = True
                    return True
        return False
    
    # comparison methods (so a list can be sorted by date)
    def __lt__(self, other):
        return self.date < other.date
    def __le__(self, other):
        return self.date <= other.date
    def __eq__(self, other):
        return self.date == other.date
    def __ne__(self, other):
        return self.date != other.date
    def __gt__(self, other):
        return self.date > other.date
    def __ge__(self, other):
        return self.date >= other.date
    def __hash__(self):
        hash_str = ''
        for i in self.original_values:
            hash_str+=i
        return hash(hash_str)
    
    
class Budget_parameters:
    
    # this is for the layered dictionary: 
            # first give 'Income' or 'Expenses'
            # third give the category
            # you will get the value back as a float
    
    # class attributes:
    # self.data_layered_dictionary (see above note)
    # self.data_dictionary
    # self.all_categories
    # self.all_income_categories
    # self.all_expense_categories
    # self.num_categories
    
    def __init__(self, budget_parameters_income_raw, budget_parameters_expenses_raw):
        self.data_layered_dictionary = {'Income':{ 'monthly':{},'weekly':{} }, 'Expenses':{ 'monthly':{},'weekly':{} } }
        self.data_dictionary = {}
        self.all_categories = []
        self.all_income_categories = []
        self.all_expense_categories = []
        
        for row in budget_parameters_income_raw:
            category, frequency, amount, notes = row
            amount = abs(float(amount))
            self.data_layered_dictionary['Income'][frequency][category]= amount
            self.data_dictionary[category] = amount
            self.all_categories.append(category)
            self.all_income_categories.append(category)
                
        for row in budget_parameters_expenses_raw:
            category, frequency, amount, notes = row
            self.data_layered_dictionary['Expenses'][frequency][category] = abs(float(amount))
            self.all_categories.append(category)
            self.all_expense_categories.append(category)
            
        self.num_categories = len(budget_parameters_income_raw) + len(budget_parameters_expenses_raw)   
    
    def get_projected_net(self):
        weekly_to_monthly = 365.25/7/12 # ew!
        weekly = sum(self.data_layered_dictionary['Income']['weekly'].values()) - sum(self.data_layered_dictionary['Expenses']['weekly'].values())
        monthly = sum(self.data_layered_dictionary['Income']['monthly'].values()) - sum(self.data_layered_dictionary['Expenses']['monthly'].values())
        projected_net_monthly = monthly + weekly*weekly_to_monthly
        projected_net_weekly = monthly/weekly_to_monthly + weekly
        return [round(projected_net_monthly,2), round(projected_net_weekly,2)] # nearest cent
        

class Budget_balancer_entry:
    #0 Timestamp	
    #1 What do you want to do?	
    #2 Category to Decrease	
    #3 Category to Increase	
    #4 Amount to Transfer	
    #5 Additional Notes
    #6 Category to Modify	
    #7 Is this an increase or a decrease?	
    #8 Should this affect the current month/week?	
    #9 Should this be one-time or permanent?	
    #10 Amount
    #11 Additional Notes
    
    ## class attributes:
    # self.row
    # self.date
    # self.is_transfer (otherwise it is an adjustment)
    # self.category
    # self.category_2
    # self.amount
    # self.amount_2
    # self.notes
    # self.affect_current
    # self.is_permanent
    
    def __init__(self, row, values):
        self.row = row
        self.date = datetime.strptime(values[0], datetime_format)
        if values[1] == 'Transfer Money':
            self.is_transfer = True
            self.category = values[2]
            self.category_2 = values[3]
            self.amount = -1 * abs(float(values[4]))
            self.amount_2 = abs(float(values[4]))
            self.notes = values[5]
            
        elif values[1] == 'Adjust Weekly/Monthly Allotmentsy':
            self.is_transfer = False
            self.category = values[6]
            if values[7] == 'Decrease':
                self.amount = -1 * abs(float(values[10]))
            else:
                self.amount = abs(float(values[10]))
            self.affect_current = (values[8] == 'Yes')
            self.is_permanent = (values[9] == 'Permanent')
            self.notes = values[11]
        else:
            pass # note sure how it would get here

def get_workbook():
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds']
#    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    creds = ServiceAccountCredentials.from_json_keyfile_name(os.path.join(os.pardir, "client_secret_secure.json"), scope) # get json file from parent directory

    client = gspread.authorize(creds)
    if run_online:
        print('Using Online Spreadsheet')
        workbook = client.open("Blackburn Budget")
    else:
        print('Using Offline Spreadsheet')
        workbook = client.open("Test Budget")
    return workbook

def read_last_sync(workbook):
    '''reads in the date and time the shreadsheet was most recently synconized'''
    worksheet = workbook.worksheet(feedpage_ws)
    last_sync_string = worksheet.cell(1,2).value
    last_sync = datetime.strptime(last_sync_string, datetime_format)
    return last_sync

def is_new_week_or_month(current_datetime, last_sync):
    '''checks to see if a new week or month has started based on the current date/time and the last sync time. Returns two booleans'''
    is_new_week = False
    is_new_month = False
    if (current_datetime - last_sync).days >=7 or current_datetime.weekday() < last_sync.weekday():
        is_new_week = True
    if current_datetime.month > last_sync.month:
        is_new_month = True
    return is_new_week, is_new_month

def read_app_input(workbook):
    '''read in all the current app input'''
    worksheet = workbook.worksheet(expense_input_ws)
    num_rows = worksheet.row_count
    app_input_raw = []
    for row_counter in range(1, num_rows+1):
        values = worksheet.row_values(row_counter)
        app_input_raw.append(values)
    # delete headers
    app_input_raw.pop(0)

    return app_input_raw

def read_bank_data(workbook):
    '''read in all the current bank data'''
    worksheet = workbook.worksheet(raw_bank_data_ws)
    num_rows = worksheet.row_count
    bank_data_raw = []
    for row_counter in range(1, num_rows+1):
        values = worksheet.row_values(row_counter)
        bank_data_raw.append(values)
    # delete headers
    bank_data_raw.pop(0)
    
    return bank_data_raw

def read_bank_unresolved(workbook):
    '''read in all the current bank data'''
    worksheet = workbook.worksheet(unresolved_items_ws)
    num_rows = worksheet.row_count
    bank_unresolved_raw = []
    for row_counter in range(1, num_rows+1):
        values = worksheet.row_values(row_counter)
        bank_unresolved_raw.append(values)
    # delete headers
    bank_unresolved_raw.pop(0)
    
    return bank_unresolved_raw


def read_budget_parameters(workbook):
    '''read in all budget parameters'''
    worksheet = workbook.worksheet(parameters_ws)
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

def read_budget_balancer_input(workbook):
    '''read in all budget balancer input'''
    worksheet = workbook.worksheet(budget_balancer_ws)
    num_rows = worksheet.row_count
    budget_balancer_input_raw = []
    for row_counter in range(1, num_rows+1):
        values = worksheet.row_values(row_counter)
        budget_balancer_input_raw.append(values)

    # delete headers
    budget_balancer_input_raw.pop(0)
    
    return budget_balancer_input_raw

def clean_input_up(app_input_raw, bank_data_raw, bank_unresolved_raw, budget_parameters_income_raw, budget_parameters_expenses_raw, budget_balancer_input_raw):
    '''cleans the input so there no funny business that happens due to a bad input value. Uses the custom classes at the top to better organize things.'''
    # return app_input, bank_data, budget_parameters, budget_balancer_input
    app_input = []
    bank_data = []
    bank_unresolved = []
    
    budget_balancer_input = []
    
    row_counter = 2 # since there is a header
    for values in app_input_raw:
        if values[0] != '': # check for blank line by looking at the first value
            app_input.append(App_entry(row_counter, values))
        row_counter +=1
        
    row_counter = 2 # since there is a header
    for values in bank_data_raw:
        if values[0] != '': # check for blank line by looking at the first value
            bank_data.append(Bank_entry(row_counter, values))
        row_counter +=1
        
    row_counter = 2 # since there is a header
    for values in bank_unresolved_raw:
        if values[0] != '': # check for blank line by looking at the first value
            bank_unresolved.append(Bank_entry(row_counter, values))
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

def create_new_month(workbook, current_datetime, last_sync, budget_parameters):
    '''makes a new worksheet titled [month year]. This is only called at the start of a month. Also fills in the category headers'''
    worksheet_names = get_all_ws_names(workbook)
    first_modified_date = current_datetime
    
    month_counter = last_sync.month
    print("last_sync:",last_sync)
#    print("last_sync.month:",last_sync.month)
    print("current_datetime:",current_datetime)
#    print("current_datetime.month:",current_datetime.month)
    year_counter = last_sync.year
    num_categories = budget_parameters.num_categories
    
    while year_counter <= current_datetime.year and month_counter <= current_datetime.month:
        current_ws_name = Month_names[month_counter] + ' ' + str(year_counter)
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
            previous_ws_name = Month_names[previous_month] + ' ' + str(previous_year)
            
            beginning_of_month = datetime(year_counter, month_counter, 1)
            first_modified_date = min(beginning_of_month,first_modified_date)
            
            if previous_ws_name in worksheet_names:
                previous_ws = workbook.worksheet(previous_ws_name)
                copy_end_of_month_values(previous_ws, worksheet, beginning_of_month)
                
            print('Created sheet named:',current_ws_name,'and added initial values.')
            # do monthly refills
            print('Doing Monthly Refills for',current_ws_name)
            monthly_refill(worksheet, beginning_of_month, budget_parameters)
        else:
            print('Did not need to create sheet named:',current_ws_name)        
        # increment the month and year counters
        month_counter += 1
        if month_counter > 12:
            month_counter -= 12
            year_counter += 1
            
    return first_modified_date

def copy_end_of_month_values(previous_ws, current_ws, beginning_of_month):
    #delete existing row if it exists:
    if current_ws.cell(3, 3) == last_month_copy_description:
        current_ws.delete_row(3)
        
    bottom_row = find_bottom_row(previous_ws)
    end_of_month_values = previous_ws.row_values(bottom_row)
    beginning_of_month_values = [0] * current_ws.col_count # default all values to zero just in case there are new categories
    
    # assign category vlues to the proper category in the next month (may not be the same column)
    for col_counter in range(7, previous_ws.col_count+1):
        new_col_number = get_column_number(current_ws, previous_ws.cell(2,col_counter).value )
        if new_col_number is not None:
            beginning_of_month_values[new_col_number - 1] =  end_of_month_values[col_counter - 1]
    
    # overwrite the first few lines
    beginning_of_month_values[0] = beginning_of_month.strftime(date_format)
    beginning_of_month_values[1] = ''
    beginning_of_month_values[2] = last_month_copy_description
    beginning_of_month_values[3] = ''
    beginning_of_month_values[4] = '0'
    beginning_of_month_values[5] = 'auto-generated'
    print('Adding the first row of data to',current_ws.title,'based on',previous_ws.title)
    
    current_ws.insert_row(beginning_of_month_values, index=3)

def get_all_ws_names(workbook):
    all_worksheets = workbook.worksheets()
    worksheet_names = []
    for worksheet in all_worksheets:
        worksheet_names.append(worksheet.title)
    return worksheet_names
            
def find_bottom_row(worksheet):
    '''finds the row closest to the bottom of the worksheet that has a non-empty cell'''
    current_row = worksheet.row_count
    while current_row>0:
        empty_row = True
        row_values = worksheet.row_values(current_row)
        for value in row_values:
            empty_row = empty_row & (value == '')
        if empty_row:
            current_row-=1
        else:
            break
    return current_row
    
def get_row_from_date(worksheet, date):
    '''return the row of the first instance of a date that is less than or equal to the given date'''
    # in case a date was passed instead of a datetime:
    date = datetime(date.year, date.month, date.day)
    row_counter = 3
    num_rows = find_bottom_row(worksheet)
    
    # find the appropriate row to insert the line:
    while row_counter <= num_rows:
        current_date_str = worksheet.cell(row_counter, 1).value
        if date <= datetime.strptime(current_date_str, date_format):
            break
        row_counter += 1
    return row_counter
    
def add_budget_line_item(worksheet, date, check_number, bank_description, app_discription, amount, category):
    '''adds a budget item to the given worksheet. Determines where to add it based on the date. 
    Assumes there is an existing line of data. Adds it AFTER items with the same date'''
    # just in case date is not a datetime:
    date = datetime(year=date.year,month=date.month,day=date.day)
    
    row_counter = 3
    num_rows = worksheet.row_count
    # find the appropriate row to insert the line:
    while row_counter <= num_rows:
        current_date_str = worksheet.cell(row_counter, 1).value
        if current_date_str == '':
            break
        if date < datetime.strptime(current_date_str, date_format):
            break
        row_counter += 1
    # copy values from the above row
    values = []
    values.append(date.strftime(date_format))
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
        return row_values.index(category) + 1 #convert from zero- to one-
    else:
        return None

def propagate_addition(workbook, init_date):
    '''starts from the given date and does all the addition from scratch to avoid errors'''
    
    # handle first worksheet
    worksheet = get_current_month_ws(workbook, init_date)
    init_row = get_row_from_date(worksheet, init_date)
    propagate_addition_ws(worksheet, init_row)
    
    # handle subsequent worksheets
    month_counter = init_date.month + 1
    year_counter = init_date.year
    date_counter = datetime(year_counter, month_counter, 1) # the first day of next month
    previous_ws = worksheet
    worksheet = get_current_month_ws(workbook, date_counter)
    while worksheet is not None:
        # update beginning of the month values
        copy_end_of_month_values(previous_ws, worksheet, date_counter)
        propagate_addition_ws(worksheet, 4) #we don't want to modify the first entry that had the values from last month
        
        # increment to the next first day of the month
        month_counter += 1
        if month_counter > 12:
            month_counter-=12
            year_counter+=1
        date_counter = datetime(year_counter, month_counter, 1)
        
        # update previous and current worksheets
        previous_ws = worksheet
        worksheet = get_current_month_ws(workbook, date_counter)

def propagate_addition_ws(worksheet, init_row):
    num_rows = find_bottom_row(worksheet)
    num_cols = worksheet.col_count
    if init_row<=3:
        init_row =4 # cannot start before the first line of data
    previous_values = worksheet.row_values(init_row-1)
    
    for row_counter in range(init_row, num_rows+1):
        values = worksheet.row_values(row_counter)
        amount = float(values[4])
        category = values[5]
        affected_column = get_column_number(worksheet, category)
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

def start_new_week(workbook, current_datetime, last_sync, budget_parameters):
    first_modified_date = current_datetime
    weekday_of_last_sync = last_sync.weekday()
    current_first_weekday = datetime(last_sync.year,last_sync.month, last_sync.day) + timedelta(days=7-weekday_of_last_sync)
    week_increment = timedelta(days=7)
    while current_first_weekday <= current_datetime:
        first_modified_date = min(first_modified_date, current_first_weekday)
        worksheet = get_current_month_ws(workbook, current_first_weekday)
        weekly_refill(worksheet, current_first_weekday, budget_parameters)
        current_first_weekday += week_increment
        
    return first_modified_date

def weekly_refill(worksheet, first_date_of_week, budget_parameters):
    '''puts weekly refills into each appropriate budget category'''
    first_day_next_month = datetime(year=first_date_of_week.year,month=first_date_of_week.month +1, day=1)
    days_left_in_month = (first_day_next_month-first_date_of_week).days
    if days_left_in_month < 7:
        partial_weekly_refill(worksheet, first_date_of_week, days_left_in_month, budget_parameters)
    else:
        # income is negative (since it hasn't happened yet)
        for category in budget_parameters.data_layered_dictionary['Income']['weekly'].keys():
            amount = budget_parameters.data_layered_dictionary['Income']['weekly'][category]
            add_budget_line_item(worksheet, first_date_of_week, '', 'Weekly Refill for '+category, '', -1*amount, category)
        # expenses are positive
        for category in budget_parameters.data_layered_dictionary['Expenses']['weekly'].keys():
            amount = budget_parameters.data_layered_dictionary['Expenses']['weekly'][category]
            add_budget_line_item(worksheet, first_date_of_week, '', 'Weekly Refill for '+category, '', amount, category)

def partial_weekly_refill(worksheet, date, num_days, budget_parameters):
    partial = num_days/7.0
    # income is negative (since it hasn't happened yet)
    for category in budget_parameters.data_layered_dictionary['Income']['weekly'].keys():
        amount = budget_parameters.data_layered_dictionary['Income']['weekly'][category]
        amount = round(amount*partial,2)
        add_budget_line_item(worksheet, date, '', 'Partial Weekly Refill for '+category, '', -1*amount, category)
    # expenses are positive
    for category in budget_parameters.data_layered_dictionary['Expenses']['weekly'].keys():
        amount = budget_parameters.data_layered_dictionary['Expenses']['weekly'][category]
        amount = round(amount*partial,2)
        add_budget_line_item(worksheet, date, '', 'Partial Weekly Refill for '+category, '', amount, category)

def monthly_refill(worksheet, first_date_of_month, budget_parameters):
    '''puts monthly refills into each appropriate budget category'''
    # income is negative (since it hasn't happened yet)
    for category in budget_parameters.data_layered_dictionary['Income']['monthly'].keys():
        amount = budget_parameters.data_layered_dictionary['Income']['monthly'][category]
        add_budget_line_item(worksheet, first_date_of_month, '', 'Monthly Refill for '+category, '', -1*amount, category)
    # expenses are positive
    for category in budget_parameters.data_layered_dictionary['Expenses']['monthly'].keys():
        amount = budget_parameters.data_layered_dictionary['Expenses']['monthly'][category]
        add_budget_line_item(worksheet, first_date_of_month, '', 'Monthly Refill for '+category, '', amount, category)
    # do partial weekly refills
    days_in_week = 7 - first_date_of_month.weekday()
    partial_weekly_refill(worksheet, first_date_of_month, days_in_week, budget_parameters)
    

def get_current_month_ws(workbook, date):
    
    worksheet_names = get_all_ws_names(workbook)
    
    current_ws_name = Month_names[date.month] + ' ' + str(date.year)
    if current_ws_name in worksheet_names:
        return workbook.worksheet(current_ws_name)
    else:
        return None

def move_and_delete_matches(workbook, app_input, bank_data, current_datetime):
    '''takes the matched bank entries and app entries, then inserts bank entries into the appropriate worksheets 
    and deletes the old information from the input sheets and moves outstanding bank entries into their own worksheet
    Also marks app entries that are left as 'Included' '''
    
    worksheet_bank_unresolved = workbook.worksheet(unresolved_items_ws)
    first_modified_date = current_datetime
    total_unresolved = 0
    
    # delete all matched app_entries
    rows_to_delete = []
    unmatched_app_entries = []
    
    worksheet_app_entries = workbook.worksheet(expense_input_ws)
    for app_entry in app_input:
        if app_entry.is_matched:
            rows_to_delete.append(app_entry.row)
        else:
            unmatched_app_entries.append(app_entry)
            worksheet_app_entries.update_cell(app_entry.row,13,'Included')
    rows_to_delete = sorted(rows_to_delete,reverse=True) # by deleting them in reverse order, we won't have any problems with the rows shifting
    
    # can't delete all unfrozen rows in worksheet, so add a fresh row if needed
    if len(rows_to_delete) == worksheet_app_entries.row_count +1:
        worksheet_app_entries.add_rows(1)
        
    for row in rows_to_delete:
        worksheet_app_entries.delete_row(row)
    
        
    # delete all bank entries
    worksheet_bank_entries = workbook.worksheet(raw_bank_data_ws)
    num_rows = worksheet_bank_entries.row_count
    for row in range(2,num_rows+1):
        worksheet_bank_entries.delete_row(2)
    worksheet_bank_entries.add_rows(1) # just to leave one blank row at the bottom so it's easier to past in data
    
    #delete all unresolved items
    num_rows = worksheet_bank_unresolved.row_count
    for row in range(2,num_rows+1):
        worksheet_bank_unresolved.delete_row(2)
    
    # input all bank_entries that are unmatched and matched into appropriate sheets
    for bank_entry in bank_data:
        if not bank_entry.is_matched:
            values = bank_entry.original_values
            worksheet_bank_unresolved.append_row(values)
            total_unresolved += 1
        else:
            worksheet = get_current_month_ws(workbook, bank_entry.date)
            add_budget_line_item(worksheet, 
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

def remove_duplicate_bank_entries(workbook, bank_data):
    '''before bank entries are assigned, we need to make sure they haven't already been inputted and 
    categorized (or put in the outstanding worksheet)'''
    # TODO: this will need to be modified to account for multi-category expenses (since the amount will not be the same)
    # My current idea to do this is to put a pre-fix on the bank description with a flag and the total amount, so we can still check it here
    
    bank_entries_to_keep = []
    entires_matched = [] # this will help handle dulplicates by tracking wich ones have been used to make a match
    for bank_entry in bank_data:
        duplicate = False
        worksheet = get_current_month_ws(workbook, bank_entry.date)
        if worksheet is None:
            bank_entries_to_keep.append(bank_entry)
            continue 
        row = get_row_from_date(worksheet, bank_entry.date)
        num_rows = find_bottom_row(worksheet)
        if row <= num_rows:
            compare_date = datetime.strptime(worksheet.cell(row,1).value, date_format).date()
            while bank_entry.date == compare_date:
                if bank_entry.description == worksheet.cell(row,3).value and str(bank_entry.amount) == worksheet.cell(row,5).value:
                    ID = [worksheet.title,row]
                    if ID not in entires_matched:
                        duplicate = True
                        entires_matched.append(ID)
                        break
                row +=1
                if row > num_rows:
                    break 
                compare_date = datetime.strptime(worksheet.cell(row,1).value, date_format).date()
        if not duplicate:
            bank_entries_to_keep.append(bank_entry)
    return bank_entries_to_keep

def merge_bank_data_and_unresolved(bank_data, bank_unresolved):
    # https://stackoverflow.com/questions/36349881/union-of-two-lists-in-python
    # both objects need a __hash__ method, and you can just return the built in hash() result
    a, b = bank_data, bank_unresolved
    na, nb = Counter(a), Counter(b)
    return list(Counter({k: max((na[k], nb[k])) for k in set(a + b)}).elements())

def delete_old_bank_data(bank_data):
    '''removes all bank data that is older than the earliest_tracking_date'''
    bank_data_to_keep = []
    for bank_entry in bank_data:
        if bank_entry.date >= earliest_tracking_date:
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
    all_worksheets = get_all_ws_names(workbook)
    date_counter = earliest_modified_date.replace(day=1,hour=0,minute=0,second=0)
    worksheet_name = Month_names[date_counter.month] + ' ' + str(date_counter.year)
    while worksheet_name in all_worksheets:
        worksheet = get_current_month_ws(workbook, date_counter)
        first_row = worksheet.row_values(3)[6:]
        last_row = worksheet.row_values(find_bottom_row(worksheet))[6:]
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
        worksheet_name = Month_names[date_counter.month] + ' ' + str(date_counter.year)
    return most_recent_net 

def str_list_to_float(strings):
    '''converts all strings to floats, https://stackoverflow.com/questions/7368789/convert-all-strings-in-a-list-to-int'''
    return list(map(float, strings)) 

def get_update_page_info(workbook, budget_parameters, unmatched_app_entries, current_datetime):
    '''returns a clean dictionary with "category:[amount remaining, next refill date, refill amount]"
    also adds in the unmatched app entries too'''
    worksheet = get_current_month_ws(workbook, current_datetime)
    next_first_weekday = datetime(current_datetime.year,current_datetime.month, current_datetime.day) + timedelta(days=7-current_datetime.weekday())
#    next_first_weekday_str = next_first_weekday.strftime(date_format)
    if current_datetime.month < 12:
        next_first_month = datetime(current_datetime.year,current_datetime.month + 1, 1)
    else:
        next_first_month = datetime(current_datetime.year + 1, 1, 1)
#    next_first_month_str = next_first_month.strftime(date_format)
    category_row = worksheet.row_values(2)[6:]
    last_row = worksheet.row_values(find_bottom_row(worksheet))[6:]
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

def update_feedback_page(workbook, current_datetime, total_unresolved, output_dictionary, most_recent_net, projected_net):
    '''clears the feedback page and updates it with the new budget category values in one step. Also recrods the "sync" time and date '''
    worksheet = workbook.worksheet(feedpage_ws)
    
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
        
        cell_list[cell_counter].value = values[1].strftime(date_format)
        cell_counter += 1
        
        cell_list[cell_counter].value = str(values[2])
        cell_counter += 1
    
    # update sync time
    print('New Sync Time is:',current_datetime.strftime(datetime_format))
    worksheet.update_cell(1,2,current_datetime.strftime(datetime_format))
    # update net values
    worksheet.update_cell(2,2,str(most_recent_net))
    
    worksheet.update_cell(3,2,str(projected_net[0]))
    worksheet.update_cell(3,4,str(projected_net[1]))
    # update expense categories
    worksheet.update_cells(cell_list)  

#def test_read_matrix(workbook):
#    ## Select a range
#    #cell_list = worksheet.range('A1:C7')
#    cell_list = workbook.worksheet('test sheet').range(2, 2, 6, 6)
#    counter = 101
#    for cell in cell_list:
#        print(cell.value)
#        cell.value = counter
#        counter+=1
#    # Update in batch
#    workbook.worksheet('test sheet').update_cells(cell_list)   

def balance_budget():
    '''manages input from the budget balancer form. Allows you to modify budget parameters permenently, retroacticely, or just one-time. Also allows you to do one-time transfers between expense categories. Also delete the rows of the form entries that are used'''
    pass

if run_online:
    try:
        main()
    except:
        pass
else:
    main()
#    main(fake_date=datetime(year=2017,month=9,day=5))

