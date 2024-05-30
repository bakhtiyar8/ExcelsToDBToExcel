1_1_CollectDataToDBmain.py

This Python code extracts data from Excel files and stores it in a SQLite database. Here's a breakdown of the code:

1. Setting Up:

Imports necessary libraries: os for file system interaction, openpyxl for reading Excel files, and sqlite3 for working with SQLite databases.
Defines the folder path containing the Excel files (folder_path).
Creates a connection to a new SQLite database named excel_data5.db.
Creates four tables in the database with specific columns to store extracted data:
TSS_date: Stores data related to TSS dates (likely related to telecommunication site maintenance).
Site_id_information: Stores information about site IDs, addresses, coordinates, etc.
excel_rf_old_data: Stores data related to old RF configurations (Radio Frequency).
RF_NEW_DATA: Stores data related to new RF configurations.
2. Processing Excel Files:

Loops through all files in the specified folder.
Checks if the file is an Excel file (ends with .xlsx).
If yes, opens the file using openpyxl.
Looks for a sheet name containing the text "SITE PARAM" (case-insensitive).
If found, assigns that sheet to a variable selected_sheet.
If not found, prints a message and continues to the next file.
3. Extracting Data:

Extracts data from specific rows and columns in the selected sheet and inserts them into corresponding tables:
Row 3 (columns A to G): Inserted into TSS_date table.
Row 8 (columns A to L): Inserted into Site_id_information table.
Row 13 (columns A to F): Inserted into excel_rf_old_data table.
Rows starting from 14 until a blank cell is encountered (columns A to F): Inserted into excel_rf_old_data table (assuming data continues for old configurations).
Extracts data for "RF NEW DATA":
Finds the starting row for "RF NEW DATA" (assuming the text appears before the data).
Loops through subsequent rows until a blank cell is encountered (columns A to L): Appends data as tuples to a list rf_new_data.
Inserts all data in rf_new_data at once into the RF_NEW_DATA table using executemany.
4. Finalizing:

Commits changes to the database (conn.commit()).
Closes the connection to the database (conn.close()).

1_2_CheckDB.py

The provided code is written in Python and utilizes the sqlite3 module to interact with a SQLite database. It aims to check for the existence of specific tables, remove spaces from a particular column in those tables, and display the contents of those tables if they exist.

Detailed Breakdown:

Establish Database Connection:

Imports the sqlite3 module.
Creates a connection to the SQLite database named excel_data5.db using sqlite3.connect().
Creates a cursor object (c) using conn.cursor().
Check for Table Existence and Remove Spaces:

For each table (excel_rf_old_data, RF_NEW_DATA, Site_id_information, TSS_date):
Checks if the table exists using c.execute() and c.fetchone().
If the table exists:
Removes spaces from the column [RF OLD DATA: Antenna / Антенна] or [RF NEW DATA: Antenna / Антенна] using c.execute() and REPLACE().
Commits the changes using conn.commit().
Retrieves all data from the table using c.execute() and c.fetchall().
If data exists:
Prints a message indicating data presence.
Iterates through the rows (row) and displays each row's contents using print(row).
Otherwise, prints a message indicating no data found.
Close Database Connection:

Closes the database connection using conn.close().
Overall Summary:

The script checks for the existence of four tables in the SQLite database.
For each existing table, it removes spaces from a specific column and displays the table's contents.
Finally, it closes the database connection.
Key Points:

Uses sqlite3 to interact with the SQLite database.
Checks for table existence, removes spaces from columns, and displays table data.
Handles each table separately.
Closes the database connection at the end.

0_2__report_to_excel.py


This code creates an Excel workbook and exports data from three tables in a SQLite database to separate sheets in the workbook. Here's a breakdown:

Import libraries:

sqlite3: Enables connecting to and interacting with the SQLite database.
openpyxl: Provides tools for creating and manipulating Excel spreadsheets.
Database connection:

Establishes a connection to the database named "excel_data5.db" using sqlite3.connect.
Creates a cursor object (c) to execute queries.
Excel workbook creation:

Initializes an empty Excel workbook using openpyxl.Workbook().
Exporting data:

Loops through three sections to export data from different tables:
RF_NEW_DATA:
Creates a sheet named "RF_NEW_DATA" in the workbook.
Retrieves column names (headers) from the table using a query.
Writes the headers to the first row of the sheet.
Fetches all rows of data from the table and writes them to subsequent rows in the sheet.
excel_rf_old_data: Similar process as above, exporting data to a sheet named "RF_OLD_DATA".
Site_Information:
Creates a sheet named "Site_Information".
Retrieves column names from both "Site_id_information" and "TSS_date" tables.
Combines the headers from both tables into a single list.
Writes the combined headers to the first row of the sheet.
Fetches all rows of data from both tables separately.
Zips the corresponding rows from each table to create combined rows.
Writes the combined rows to the sheet.
Saving and closing:

Saves the workbook as "TSSR report data.xlsx".
Closes the connection to the database using conn.close().
Success message:

Prints a message indicating successful report export.
