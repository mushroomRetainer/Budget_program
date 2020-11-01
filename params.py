
from datetime import datetime, timedelta

class Params:
    # adjustable parameters
    def __init__(self):
        self.run_online = True
        
        self.Month_names = ['month0','Jan','Feb','March','April','May','June','July','Aug','Sept','Oct','Nov','Dec']
        
        self.datetime_format = '%m/%d/%Y %H:%M:%S'
        self.date_format = '%m/%d/%Y' # hopefully the same as BECU to cut down on formatting
        self.last_month_copy_description = 'Copy from End of Last Month'
        self.date_match_tolerance = timedelta(days=3)
        self.earliest_tracking_date = datetime(year=2020,month=1,day=1).date() # anything before this will be deleted
        # other sheets are created and named dynamically based on month/year
        
        # inpput worksheet names
        self.feedpage_ws = 'Current Budget'
        self.expense_input_ws = 'Input From Expense Log'
        self.balancer_input_ws = 'Input From Budget Balancer'
        self.parameters_ws = 'Budget Parameters'
        self.budget_balancer_ws = 'Input From Budget Balancer'
        self.auto_transactions_ws = 'Automatic transactions'
        self.raw_bank_data_ws = 'raw bank acount data'
        self.unresolved_items_ws = 'unresolved bank items'