
from datetime import datetime, timedelta
import itertools
from collections import Counter # needed for complex merging to preserve duplicates in individual lists, but not in both

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
    
    def __init__(self, row, values, params):
        self.row = row # keep in mind this is the spreadsheet row, so it is 1-indexed
        self.original_values = values
        self.params = params
        
        if values[3] != '':
            self.date = datetime.strptime(values[3], self.params.date_format).date()
        elif values[0] != '':
            self.date = datetime.strptime(values[0], self.params.datetime_format).date()
        
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
    
    def __init__(self, row, values, params):
        self.params = params
        self.row = row
        self.original_values = values
        self.date = datetime.strptime(values[0], self.params.date_format).date()
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
        elif abs(self.date - app_entry.date) <= self.params.date_match_tolerance and self.amount == app_entry.amount:
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
            if app_entry.is_combined and abs(self.date - app_entry.date) <= self.params.date_match_tolerance:
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
    
    def __init__(self, row, values, params):
        self.row = row
        self.params = params
        self.date = datetime.strptime(values[0], self.params.datetime_format)
        if values[1] == 'Transfer Money':
            self.is_transfer = True
            self.category = values[2]
            self.category_2 = values[3]
            self.amount = -1 * abs(float(values[4]))
            self.amount_2 = abs(float(values[4]))
            self.notes = values[5]
            
        elif values[1] == 'Adjust Weekly/Monthly Allotments':
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