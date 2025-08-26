#region imports
from sys import prefix
import pyodbc
import json
from openpyxl import load_workbook
from typing import Any
import os
#endregion


#region functions

def read_excel(excel_path: str) -> list[Any]:
    data:list[Any]= [] 
    if "~$" not in excel_path:
        wb = load_workbook(excel_path)
        sheets = wb.sheetnames
        for sheet in sheets:
            data.append(wb[sheet]) 
    return data

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
    
    #reguib sql utils
    def table_query_builder(self, table_name:str, cols:list[str], data_types:list[str]) -> str:

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
        for data in self.excel_datas:
            sample:Any = data["I3"].value #type:ignore 
            run:Any = data["N3"].value #type:ignore
            table_name:str = f"cvd_{sample}_{run}"
            data_holder = self.cursor.execute("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'")
            self.tables = [x[2] for x in data_holder]
            col_name_locations:list[str] = ["A3", "H3", "M3", "A5", "H5"]
            
            col_names:list[str] = []
            col_data_types:list[str] = ["VARCHAR(255)", "VARCHAR(255)", "VARCHAR(255)", "VARCHAR(255)", "VARCHAR(255)"]
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
            datas.append(read_excel(os.path.join("datasheets",file)))
    temp = sql_data_handler("config.json", datas[0])
    
    
    temp.connect()
    
    temp.build_table()
    
    temp.close()
#end region