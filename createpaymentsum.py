# Reads excel file, keeps specified columns, drops duplicate records, 
# writes to a file, and formats the header and columns
import pandas as pd
import numpy as np
import paymentfunctions as f
# define variables and functions
new_data = "MLEPCVBSPayments2018_Data.xlsx"
working_data = "MLEPCVBSPayments2018.xlsx"
payment_summary = "MLEPCVBSPaymentBalances2018.xlsx"
sheet_name = 'Worksheet'
master_sheet = 'Outstanding Payments'

# read the data file - new_data
try:
    new_df = pd.read_excel(new_data, sheet_name)
except FileNotFoundError:
    # the file does not exist
    print('The file does not exist. Download the child information file and run again.')
    exit(0)
else:
    # the file exists - continue
    pass

# load lines greater than given Time
# keep only records since the last pull
last_pull = '4/12/2018  7:00:00 AM'
new_df = new_df[(new_df['Time'] > last_pull)]

# add column for payment balance
new_df['Balance'] = new_df['Total Cost']

# Keep only columns relevant to evaluating payment balances
keep_col = ['Balance','Parent or Volunteer Last Name','Parent or Volunteer First Name','Payment Status','Email','Total Cost','Check Amt']
new_df = new_df[keep_col]
new_df = new_df.rename(index=str, columns={"Parent or Volunteer Last Name": "Last Name", "Parent or Volunteer First Name": "First Name"})

# add column for payment balance
for rw in range(0,len(new_df)):
    if not (np.isnan(new_df.iat[rw, 6])):
            new_df.iat[rw, 0] = new_df.iat[rw, 5] - new_df.iat[rw, 6]
    if new_df.iat[rw, 3] == 'Completed':
            new_df.iat[rw, 0] = 0
'''
# add column for payment balance
for col in range(0,len(new_df.columns)):
       for rw in range(0,len(new_df)):
           if col == 0:
               if not (np.isnan(new_df.iat[rw, 6])):
                   new_df.iat[rw, col] = new_df.iat[rw, 5] - new_df.iat[rw, 6]
               if new_df.iat[rw, 3] == 'Completed':
                   new_df.iat[rw, col] = 0
  '''

# keep records with a positive or negative balance
new_df = new_df.loc[new_df['Balance'] != 0]

# convert all names to proper case (initial caps)
# strip leading and trailing spaces
str_to_format = ['First Name','Last Name']
new_df = f.format_str(new_df, str_to_format)

# sort dataframe on last grade completed, child last name, child first name
new_df = new_df.sort_values(by=['Payment Status','Last Name','First Name'])
# create the file and write the results to the file
writer = pd.ExcelWriter(payment_summary)
workbook = writer.book
new_df.to_excel(writer, master_sheet, index=False, startrow=1, header=False)
writer = f.format_info(new_df, writer, master_sheet)
writer.save()
