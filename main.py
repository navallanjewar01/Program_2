# -*- coding: utf-8 -*-
"""
Created on Tue May  7 13:07:55 2024

@author: Pyadmin
"""


import pandas as pd
import re

# Function to parse dates with different formats
def parse_date(date_str):
    """
    Parse dates from strings with different formats.

    Args:
        date_str (str): Date string to parse.

    Returns:
        pd.Timestamp: Parsed datetime object.
    """
    # Regular expressions to match different date formats
    patterns = [
        r'(?P<month>\d{1,2})/(?P<day>\d{1,2})/(?P<year>\d{4})',
        r'(?P<month>\d{1,2})/(?P<day>\d{1,2})/(?P<year>\d{2})'
    ]
    for pattern in patterns:
        match = re.match(pattern, date_str)
        if match:
            year = match.group('year')
            # Convert 2-digit year to 4-digit year
            if len(year) == 2:
                year = '20' + year
            return pd.to_datetime(match.group('month') + '/' + match.group('day') + '/' + year, format='%m/%d/%Y')
    return pd.NaT  # Return NaT (Not a Time) if no valid date format is found

# Read the Excel file
input_file = r"C:\Users\Pyadmin\Program_2\SampleData.xlsx"
output_file = r"C:\Users\Pyadmin\Program_2\output.xlsx"
df = pd.read_excel(input_file)

# Replace empty cells in "SalesMan" column with your name
df['SalesMan'].fillna("Naval", inplace=True)

# Convert values in "OrderDate" column to datetime objects
df['OrderDate'] = df['OrderDate'].apply(parse_date)

# Replace NaT (Not a Time) values with empty string
df['OrderDate'] = df['OrderDate'].apply(lambda x: x.strftime("%d/%m/20%y") if not pd.isnull(x) else "")

# Write the updated data to a new Excel file
df.to_excel(output_file, index=False)

print("File has been created successfully with updated data.")

