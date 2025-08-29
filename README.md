# cvd_sql
cvd sql database builder as well as other cvd sql utils


This code is a Python script designed to automate the extraction of data from Excel files and insert it into a SQL Server database. It uses several libraries: openpyxl for reading Excel files, pyodbc for database connectivity, and standard libraries like json, os, and shutil for configuration and file management.

The read_excel function loads an Excel workbook and returns a list of worksheet objects, skipping temporary files (those with ~$ in their name). This function is used to gather data from all Excel files in a specified directory.

The core logic is encapsulated in the sql_data_handler class. When instantiated, this class reads configuration details (such as database connection parameters) from a JSON file, and stores the Excel data and file paths. The class provides several methods for processing Excel data: it can generate complex column names based on cell ranges and headers, build SQL queries for creating tables, and manage the addition of columns to those tables. The methods section_to_cols, gen_col_names, and complex_addy work together to extract and format column names from structured regions of the Excel sheets, handling special cases like comments and special characters.

The build_table method creates a new SQL table for each Excel sheet, using specific cell values to generate the table name and column names. If the table does not already exist, it is created with a basic set of columns. The build_cols method then adds additional columns to the table, based on the dynamically generated column names from the Excel data. The execute method orchestrates the process for all loaded Excel files, iterating through each sheet, generating columns, creating tables, extracting data, and writing it to the database. After processing, files are moved to a "processed" directory.

Database connections are managed with the connect and close methods, which open and close the connection using the parameters loaded from the configuration file. The script’s entry point (if __name__ == "__main__":) walks through the "to_process" directory, reads all Excel files, initializes the sql_data_handler, connects to the database, processes the files, and then closes the connection.

Overall, this script provides a flexible and automated way to import structured data from Excel files into a SQL Server database, with configurable connection settings and robust handling of complex Excel layouts.



## usage


```bash
excel2sql <your command here>
```

commands = builddb,write,nomove

write-outputs all data in files in processed and then moves files to to_processed
builddb-should only be ran on the sql computer, builds the db from a sample execel sheet left in the data folder
must have one file into to run
nomove-will not move the files
below is an example to both build the db and get the data
```bash
excel2sql builddb write
```
order does not matter the code will always build the db first then write then move or not move files.

please look at the config_def.json and enter the info required into it