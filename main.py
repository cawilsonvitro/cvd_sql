#region imports
import pyodbc
import json
from openpyxl import load_workbook
from typing import Any
import os
#endregion


#region functions

def read_excel(excel_path: str) -> list[Any]:
    """
    Reads an Excel file and returns a list of worksheet objects.

    Args:
        excel_path (str): The file path to the Excel workbook.

    Returns:
        list[Any]: A list containing worksheet objects from the workbook. 
                   If the file path contains '~$', an empty list is returned.

    Note:
        The function skips files that appear to be temporary Excel files (containing '~$').
    """
    data:list[Any]= [] 
    if "~$" not in excel_path:
        wb = load_workbook(excel_path)
        sheets = wb.sheetnames
        for sheet in sheets:
            data.append(wb[sheet]) 
    return data

#endregion
#region Class

class sql_data_handler():
    #region setup
    def __init__(self, config: str, datas: list[list[Any]]) -> None:
        self.config_path = config
        self.config: dict[str,dict[str,str]] = {}
        self.dbconfig: dict[str,str] = {}
        self.cnxn = None
        self.cursor: pyodbc.Cursor 
        self.server = ""
        self.alphabet = [chr(i) for i in range(ord('A'), ord('Z')+1)]
        self.database = ""
        self.username = ""
        self.password = ""
        self.driver = "{ODBC Driver 18 for SQL Server}"
        self.trust_server_certificate = "yes"
        self.encrypt = "yes"
        self.timeout = 30
        self.excel_datas = datas
        
        self.read_config()
    
    def read_config(self) -> None:
        with open(self.config_path, 'r') as f:
            self.config = json.load(f)
        self.dbconfig = self.config.get("Database_Config", {})
        self.server = self.dbconfig.get("host","localhost")
        self.database = self.dbconfig.get("db","cvd_test")
        self.driver = self.dbconfig.get("driver","{ODBC Driver 17 for SQL Server}")
        self.username = self.dbconfig.get("username", "")
        if self.username == "": self.username = input("No username in config please enter:")
        self.password = self.dbconfig.get("password", "")
        if self.password == "": self.password = input("No password in config please enter:")
    #endregion
    
    #reguib sql and excel utils
    
    def complex_addy(self, alpha_range:list[str], number_range:list[list[int]], header:str) -> list[tuple[str,list[str]]] :
        """
        Generates a list of address strings by combining each letter in `alpha_range` with a range of numbers specified in `number_range`,
        and returns them as a list of tuples with the provided header.
        Args:
            alpha_range (list[str]): A list of string prefixes (typically letters) to be combined with numbers.
            number_range (list[list[int]]): A list of [start, end] pairs, where each pair defines the inclusive range of numbers to append to the corresponding letter in `alpha_range`.
            header (str): A string to be used as the header in the returned tuple.
        Returns:
            list[tuple[str, list[str]]]: A list containing a single tuple, where the first element is the header and the second element is the list of generated address strings.
        """
        i = 0 
        addys: list[str] = []
        for letter in alpha_range:
            for number in range(number_range[i][0], number_range[i][1] + 1):
                addy = letter + str(number)
                addys.append(addy)
            i += 1

        complex_col_name_locations: list[tuple[str,list[str]]] = [(header, addys)]
        

        return complex_col_name_locations
    
    def table_query_builder(self, table_name:str, cols:list[str], data_types:list[str]) -> str:
        """
        Builds a SQL CREATE TABLE query string for the specified table name, columns, and data types.
        Args:
            table_name (str): The name of the table to be created.
            cols (list[str]): A list of column names for the table.
            data_types (list[str]): A list of data types corresponding to each column.
        Returns:
            str: A SQL query string for creating the table with the specified columns and data types.
        Note:
            Only columns whose data type string contains both '(' and ')' will be included in the query.
        """

        query_pre:str = f"CREATE TABLE {table_name} ("
        temp:str = ""
        query_as_list:list[str] = []
        i = 0 
        for col in cols:
            if "(" in data_types[i] and ")" in data_types[i]:temp = f"\"{col}\" {data_types[i]}"
            query_as_list.append(temp)
            i += 1
            
            
        query_col = ",".join(query_as_list)
        query = f"{query_pre} {query_col})"

        return query

    
    #endregion
    
    #region excel to sql
    
    def build_table(self):
        """
        Builds and creates a SQL table for each dataset in self.excel_datas based on extracted sample and run values.
        For each dataset:
            - Constructs a table name using sample and run identifiers.
            - Checks if the table already exists in the database.
            - Extracts column names from specified Excel cell locations.
            - Dynamically generates complex column names for pre-coating check data.
            - Builds a CREATE TABLE SQL query with the extracted column names and predefined data types.
            - Executes the query to create the table if it does not already exist.
        Side Effects:
            - Updates self.tables with the list of existing tables.
            - Prints the names of complex columns and confirmation messages upon table creation.
            - Commits changes to the database.
        Assumes:
            - self.excel_datas is a list of dictionaries mapping cell addresses to cell objects.
            - self.cursor is a database cursor object with execute and commit methods.
            - self.complex_addy is a method that generates complex column address mappings.
            - self.table_query_builder is a method that builds a CREATE TABLE SQL query.
        """
        for data in self.excel_datas:
            sample:Any = data["I3"].value #type:ignore 
            run:Any = data["N3"].value #type:ignore
            table_name:str = f"cvd_{sample}_{run}"
            data_holder = self.cursor.execute("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'")
            self.tables = [x[2] for x in data_holder]
            col_name_locations:list[str] = ["A3", "H3", "M3", "A5", "H5"]
            col_data_types:list[str] = ["VARCHAR(255)", "VARCHAR(255)", "VARCHAR(255)", "VARCHAR(255)", "VARCHAR(255)"]
            
            #building complex column names
            
            #for pre-coatingcheck list
            # pre_coating_check#
            precoat_alpha_range: list[str] = ["A", "C", "E", "G"]#[self.alphabet[i] for i in range(self.alphabet.index("A"), self.alphabet.index("G")+1)]
            precoat_number_range:list[list[int]] = [[9,14]] * (len(precoat_alpha_range) - 1)
            precoat_number_range.append([9,11])

            complex_col_name_locations: list[tuple[str,list[str]]] = self.complex_addy(precoat_alpha_range, precoat_number_range, "A7")
            for complex_col in complex_col_name_locations:
                prefix:str = str(data[complex_col[0]].value)
                for col in complex_col[1]:
                    name = f"{prefix}:{data[col].value}"
                    print(name)
                pass
            col_names:list[str] = []
           
            for loc in col_name_locations:
                col_names.append(data[loc].value.replace(":",""))
            table_query = self.table_query_builder(table_name, col_names, col_data_types)
            
            if table_name not in self.tables:
                self.cursor.execute(table_query)
                self.cursor.commit()
                print(f"Table {table_name} created")


    
    #endregion
    
    #region sql server connections
    def connect(self) -> None:
        self.sql = pyodbc.connect(
            host=self.server,
            user=self.username,
            password=self.password,
            database=self.database,
            driver=self.driver,
            autocommit=True
        )
        self.cursor = self.sql.cursor()

    def close(self) -> None:
        if self.sql:
            self.sql.close()
    #endregion
    
#region entry point
if __name__ == "__main__":

    
    datas:list[list[Any]] = []
    for _,_,files in os.walk("datasheets"):
        for file in files:
            if ".xlsx" in file: datas.append(read_excel(os.path.join("datasheets",file)))
    temp = sql_data_handler("config.json", datas[0])
    
    
    temp.connect()
    
    temp.build_table()
    
    temp.close()
#end region