import pandas as pd

# Define file paths
input_file = 'data_files.xlsx'
output_file = 'output.xlsx'

# Load the sheets from the "data_files.xlsx" file
sheet1 = pd.read_excel(input_file, sheet_name='Sheet1', engine='openpyxl')
sheet2 = pd.read_excel(input_file, sheet_name='Sheet2', engine='openpyxl')

# Load the existing sheets from the "output.xlsx" file
# Note: If the sheets don't exist yet, this will create new empty DataFrames
try:
    output_sheet_a = pd.read_excel(output_file, sheet_name='Sheet_A', engine='openpyxl')
except ValueError:
    output_sheet_a = pd.DataFrame()

try:
    output_sheet_b = pd.read_excel(output_file, sheet_name='Sheet_B', engine='openpyxl')
except ValueError:
    output_sheet_b = pd.DataFrame()

# Copy the "Reference Number" column from Sheet1 to Sheet_A
if 'Reference Number' in sheet1.columns:
    reference_number = sheet1[['Reference Number']]
    output_sheet_a = pd.concat([output_sheet_a, reference_number], axis=1)

# Remove the "Remove Me" column from Sheet2 and prepare the cleaned data
if 'Remove Me' in sheet2.columns:
    sheet2_cleaned = sheet2.drop(columns=['Remove Me'])
else:
    sheet2_cleaned = sheet2

# Save the updated data to "output.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    # Save the updated "Sheet_A"
    output_sheet_a.to_excel(writer, sheet_name='Sheet_A', index=False)
    
    # Save the cleaned "Sheet_B"
    sheet2_cleaned.to_excel(writer, sheet_name='Sheet_B', index=False)
