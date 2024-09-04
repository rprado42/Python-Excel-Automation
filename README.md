# Python-Excel-Automation

This is a short projet to automate a Excel routine.


To accomplish the tasks you described using Excel, follow these steps:

### 1. Copy the "External reference number" from Tabela3 to Tabela1

1. **Open Excel** and load all three files: `close.xlsm`, `obs.xlsx`, and `snow.xlsx`.
2. Go to the **"snow.xlsx"** workbook.
3. Select the column with the header "External reference number" in Tabela3.
4. **Copy** the selected column (`Ctrl + C`).
5. Switch to the **"close.xlsm"** workbook.
6. Navigate to the **"INC MARS"** sheet.
7. **Paste** the copied data into the desired column (you may choose the appropriate location where this data should be pasted, such as a specific column in the "INC MARS" sheet).

### 2. Delete the "Number" column and copy the remaining data from Tabela2 to Tabela1

1. **Switch to the "obs.xlsx"** workbook.
2. Locate the column with the header "Number" in Tabela2.
3. **Select** the entire column containing the "Number" data.
4. **Delete** the selected column (`Right-click` on the column header and select `Delete`).

5. **Select** the remaining data (excluding the deleted "Number" column) by clicking and dragging to highlight the data you want to copy.
6. **Copy** the selected data (`Ctrl + C`).

7. **Switch** to the **"close.xlsm"** workbook.
8. Go to the **"OBS"** sheet.
9. **Paste** the copied data into the desired location (such as starting from the top-left corner of the target area in the "OBS" sheet).

### Summary

- **From `snow.xlsx`**: Copy the "External reference number" and paste it into the "INC MARS" sheet in `close.xlsm`.
- **From `obs.xlsx`**: Delete the "Number" column and copy the remaining content, then paste it into the "OBS" sheet in `close.xlsm`.

Make sure to save your work regularly and verify that all data has been transferred correctly.
