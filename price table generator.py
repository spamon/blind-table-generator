#THIS SCRIPT IS USED TO GENERATE TABLES INETBWEEN TABLES FOR BLIND PRICES, IE, HAVE 2 TABLES IN DIFFERENT SHEETS
#THEN THIS SCRIPT WILL GENERATE X AMOUNT OF TABLES REQUIRED SPLIT EVENALLY ACROSS THE PRICES
import pandas as pd

# Read the two marker tables from Excel
table1 = pd.read_excel('vertical_blinds_a.xlsx', sheet_name='Table1', header=None)
table2 = pd.read_excel('vertical_blinds_a.xlsx', sheet_name='Table2', header=None)

# Calculate the price difference and the number of intervals
price_difference = table2.values - table1.values
num_intervals = 3

# Calculate the increment value
increment = price_difference / (num_intervals + 1)

# Create the intermediate tables
tables = [table1]
for i in range(num_intervals):
    # Calculate the prices for each intermediate table
    prices = table1.values + increment * (i + 1)
    
    # Create a new table with the same structure as the marker tables
    intermediate_table = pd.DataFrame(prices)
    
    # Append the intermediate table to the list of tables
    tables.append(intermediate_table)
tables.append(table2)

# Write the tables to a new Excel file
with pd.ExcelWriter('output.xlsx') as writer:
    sheet_names = ['Table1', 'Intermediate1', 'Intermediate2', 'Intermediate3', 'Table2']
    for sheet_name, table in zip(sheet_names, tables):
        table.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
