import csv
from os import pardir, sep
import pandas as pd
import openpyxl

# Read Excel File
read_file = pd.read_excel("Orders-With Nulls.xlsx")
# Create CSV file with all data
read_file.to_csv("excel_to_csv.csv" , index=None , header=True)

# Read Specific column
col_list = ['Order Date' , 'Sales']
df = pd.read_csv("excel_to_csv.csv" , usecols=col_list)

# Sorting Date (Recent first)

df['Order Date'] = pd.to_datetime(df['Order Date'])
df.sort_values(by=['Order Date'] , inplace=True , ascending=False)

# Print Columns

print(df)

# Open existing Result CSV file 
try:
    with open("result.csv" , "r") as new_csv:
        if new_csv:
            readcsv = pd.read_csv("result.csv" , usecols=col_list)
            print(readcsv) 
            
# Create new CSV file with only two column if file not exist
# If result.csv file is not exist then this block is run and create new csv file which have only two columns Order date and sales 
# with the decending sorting (recent first) 
except:
    df.to_csv("result.csv" , sep=',' , )
    df.sort_values(by="Order Date")

