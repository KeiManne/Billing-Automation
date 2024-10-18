#This is a program to output a csv file with an up to date billing sheet given two csv files
import pandas as pd
import numpy

# Read in the completed jobs sheet and the WIP sheet
jobs_df = pd.read_excel('jobs.xlsx', usecols=['Client Ref', 'Client Name', 'Description', 'Staff', 'Job Notes', 'End Date'], header=1)
wip_df = pd.read_excel('wip.xlsx', usecols=['Client', 'Amount'], header=10)

#Format wip sheet to be same format as the jobs sheet
wip_df = wip_df.dropna(subset=['Amount'])
wip_df[['Client Ref', 'Client Name']] = wip_df['Client'].str.split(' -- ', expand=True)
wip_df = wip_df.drop('Client', axis=1)
wip_df = wip_df[['Client Ref', 'Client Name', 'Amount']]

# Merge the two dataframes on the 'Client Reference Code' column
billingsheet_df = pd.merge(jobs_df, wip_df, on='Client Ref')
billingsheet_df = billingsheet_df.drop('Client Name_x', axis=1)
billingsheet_df = billingsheet_df[['Client Ref', 'Client Name_y', 'Description', 'Staff', 'End Date', 'Job Notes', 'Amount']]
billingsheet_df.rename({'Client Name_y': 'Client Name'}, axis=1, inplace=True)

# Define a function to merge the 'Description' column and staff column and notes for a client code
def merge_desc(series):
    return ', '.join(series.dropna().unique())

# Define a function to merge the 'Staff', 'End Date', 'Job Notes' columns
def merge_other(series):
    return ', '.join(series.dropna().unique())

# Use the agg() function to apply the merge functions to the appropriate columns
groupby_df = billingsheet_df.groupby('Client Ref')
mergedbilling_df = groupby_df.agg({
    'Client Name': 'first',
    'Description': merge_desc,
    'Staff': merge_other,
    #'End Date': merge_other,
    'Job Notes': merge_other,
    'Amount': 'first'
}).reset_index()

#Final Report Formatting
#mergedbilling_df = mergedbilling_df[['Client Ref', 'Client Name', 'Description', 'Staff', 'Amount', 'End Date', 'Job Notes']]
#mergedbilling_df.rename({'Amount': 'WIP Amount'}, axis=1, inplace=True)


print(mergedbilling_df.head())
print(mergedbilling_df.tail())

# Output the merged dataframe to a new Excel sheet
mergedbilling_df.to_excel('Billing_Sheet.xlsx', index=False)