# Python-Excel-Automation

This is a short projet to automate a Excel routine.

For this project will use:

- `data_files.xlsx` for the file containing the sheets to manipulate.
  - `Sheet1` will contain the data to be copied.
  - `Sheet2` will contain the column to be deleted.
- `output.xlsx` for the output file.
  - `Sheet_A` will be the destination for the copied data.
  - `Sheet_B` will be the destination for the cleaned data.

### Updated Steps for Automation

#### 1. **Prepare the Environment**

   - **Install Required Libraries**:
     Ensure `pandas` and `openpyxl` libraries are installed:
     
     ```bash
     pip install pandas openpyxl
     ```

#### 2. **Setup the Files**

   - **Identify the Files and Sheets**:
     - `data_files.xlsx`:
       - `Sheet1` contains the column to be copied.
       - `Sheet2` contains the column to be deleted.
     - `output.xlsx`:
       - `Sheet_A` is where the column will be copied.
       - `Sheet_B` is where the cleaned data will be saved.

#### 3. **Load the Data**

   - **Read the Excel Files**:
     Use `pandas` to read the data from the specified sheets:
     
     ```python
     import pandas as pd

     # Load the sheets from the "data_files.xlsx" file
     sheet1 = pd.read_excel('data_files.xlsx', sheet_name='Sheet1', engine='openpyxl')
     sheet2 = pd.read_excel('data_files.xlsx', sheet_name='Sheet2', engine='openpyxl')

     # Load the sheets from the "output.xlsx" file
     output_sheet_a = pd.read_excel('output.xlsx', sheet_name='Sheet_A', engine='openpyxl')
     output_sheet_b = pd.read_excel('output.xlsx', sheet_name='Sheet_B', engine='openpyxl')
     ```

#### 4. **Manipulate the Data**

   - **Copy a Column from `Sheet1` to `Sheet_A`**:
     Assume the column to copy is labeled "Reference Number". Extract this column from `Sheet1` and add it to `Sheet_A`:
     
     ```python
     reference_number = sheet1[['Reference Number']]
     output_sheet_a = pd.concat([output_sheet_a, reference_number], axis=1)
     ```

   - **Delete a Column and Copy Remaining Data from `Sheet2` to `Sheet_B`**:
     Remove a column labeled "Remove Me" from `Sheet2` and prepare the remaining data to be copied:
     
     ```python
     sheet2_cleaned = sheet2.drop(columns=['Remove Me'])
     ```

#### 5. **Save the Changes**

   - **Write the Updated Data**:
     Use `pd.ExcelWriter` to save the modified data to `output.xlsx`. Ensure `Sheet_A` and `Sheet_B` are updated or created:
     
     ```python
     with pd.ExcelWriter('output.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
         # Save the updated "Sheet_A"
         output_sheet_a.to_excel(writer, sheet_name='Sheet_A', index=False)
         
         # Save the cleaned "Sheet_B"
         sheet2_cleaned.to_excel(writer, sheet_name='Sheet_B', index=False)
     ```

#### 6. **Execute the Script**

   - **Run the Script**:
     Save the code in a Python file, for example, `automate_excel.py`, and run it from your terminal or command prompt:
     
     ```bash
     python automate_excel.py
     ```

### Summary of Steps with Updated Names

1. **Prepare the Environment**: Install `pandas` and `openpyxl`.
2. **Setup the Files**: Identify the files and sheets with new names.
3. **Load Data**: Read data from `data_files.xlsx` and `output.xlsx`.
4. **Manipulate Data**: Copy the "Reference Number" column and remove the "Remove Me" column.
5. **Save Changes**: Update `output.xlsx` with the modified data.
6. **Execute the Script**: Run the Python script to apply the changes.
