
# python imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import itertools
from collections import Counter # needed for complex merging to preserve duplicates in individual lists, but not in both
import os

# adjustable parameters
run_online = True

Month_names = ['month0','Jan','Feb','March','April','May','June','July','Aug','Sept','Oct','Nov','Dec']

datetime_format = '%m/%d/%Y %H:%M:%S'
date_format = '%m/%d/%Y' # hopefully the same as BECU to cut down on formatting
last_month_copy_description = 'Copy from End of Last Month'
date_match_tolerance = timedelta(days=3)
earliest_tracking_date = datetime(year=2017,month=10,day=1).date() # anything before this will be deleted
# other sheets are created and named dynamically based on month/year

# inpput worksheet names
feedpage_ws = 'Current Budget'
expense_input_ws = 'Input From Expense Log'
balancer_input_ws = 'Input From Budget Balancer'
parameters_ws = 'Budget Parameters'
budget_balancer_ws = 'Input From Budget Balancer'
auto_transactions_ws = 'Automatic transactions'
raw_bank_data_ws = 'raw bank acount data'
unresolved_items_ws = 'unresolved bank items'