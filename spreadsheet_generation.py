# Import pandas for data manipulation
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

# Read the CSV files
session_counts = pd.read_csv('DataAnalyst_Ecom_data_sessionCounts.csv')
add_to_cart = pd.read_csv('DataAnalyst_Ecom_data_addsToCart.csv')

# General strategy: Form 2 data frames for each of the pages of the requested xlsx workbook

# Start with  Month * Device aggregation data frame

# Form a month column in the session_counts data frame
# Since the date is formatted as a string, simply find
# the first "/" and take the preceding number
# Then convert to an integer, so it sorts numerically
session_counts['month'] = session_counts['dim_date'].apply(lambda x: int(x[:x.find("/")]))

# Test this by validating some data
# This also confirms that date is in mm/dd/yyyy format
print(session_counts[session_counts['month'] == 11])
print(session_counts[session_counts['month'] == 6])

# Aggregate by month and device
xlsx_page1 = session_counts.groupby(by = ['month', 'dim_deviceCategory']).sum()

# Now simply add the the ECR metric
xlsx_page1['ECR'] = xlsx_page1['transactions'] / xlsx_page1['sessions']

# Move onto the month-over-month page

# I'm going to pretend I don't know the last 2 months of data
# by peaking inside of the files so it's going to be determined
# programmatically

# A really easy way to do this is by combining the year and month into
# a string and sorting alphabetically

# Extract the year
session_counts['year'] = session_counts['dim_date'].apply(lambda x: x[-2:])
# Combine with month
# But first convert month back into a string
session_counts['month'] = session_counts['month'].apply(lambda x: str(x))
session_counts['date_sort'] = session_counts['year'] + session_counts['month']

# Do the same thing on the add_to_cart data
# The month/year here are stored as numbers, so they'll need to be converted to strings
add_to_cart['dim_month'] = add_to_cart['dim_month'].apply(lambda x: str(x))
# Year is as 4 digit number, so convert to string and take the last two digits, to be consistent
add_to_cart['dim_year'] = add_to_cart['dim_year'].apply(lambda x: str(x)[-2:])
# combine the two
add_to_cart['date_sort'] = add_to_cart['dim_year'] + add_to_cart['dim_month']

# Get the two most recent year/months
# I'm assuming the data will both have the same most recent month/year
date_values = np.sort(session_counts['date_sort'].unique())[-2:]

# Filter out the two most recent months
session_counts = session_counts[(session_counts['date_sort'] == date_values[0]) | (session_counts['date_sort'] == date_values[1])]
add_to_cart = add_to_cart[(add_to_cart['date_sort'] == date_values[0]) | (add_to_cart['date_sort'] == date_values[1])]

# Group the session counts data by month
xlsx_page2 = session_counts.groupby(by = ['month']).sum()

# Add the the ECR metric
xlsx_page2['ECR'] = xlsx_page2['transactions'] / xlsx_page2['sessions']

# Add the add_to_carts by merging the data frames
xlsx_page2 = xlsx_page2.merge(add_to_cart, left_on = 'month', right_on = 'dim_month')

# Remove the date_sort and dim_year columns
xlsx_page2 = xlsx_page2.drop(['date_sort', 'dim_year'], axis = 1)

# It will make relative changes much easier (and readable)
# if columns/rows flip
# Take the transpose of the data
xlsx_page2 = xlsx_page2.T

# Promote the dim_month to headers
month_row = xlsx_page2.iloc[4]
month1 = 'month ' + month_row[0]
month2 = 'month ' + month_row[1]
xlsx_page2.columns = [month1, month2]

# The row with the month is no longer needed
xlsx_page2 = xlsx_page2.drop(index = ['dim_month'])

# finally add absolute and relative difference
# Recall that the names of columns won't be known ahead of time
# (because it's based on the number of the most recent months)
# Manipulation should be based on column indices instead
xlsx_page2['difference'] = xlsx_page2[xlsx_page2.columns[1]] - xlsx_page2[xlsx_page2.columns[0]]
xlsx_page2['% change'] = xlsx_page2['difference'] / xlsx_page2[xlsx_page2.columns[0]]

# Finally export this all to an excel file
with pd.ExcelWriter('output.xlsx') as writer:
    xlsx_page1.to_excel(writer, sheet_name='Device Month Summary')
    xlsx_page2.to_excel(writer, sheet_name='Month over Month Comparison')