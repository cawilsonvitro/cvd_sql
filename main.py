#region imports
import pyodbc
import json
from openpyxl import load_workbook
from typing import Any
#endregion


#region functions

def read_excel(excel_path: str) -> list[Any]:
    wb = load_workbook(excel_path)
    sheets = wb.sheetnames
    data:list[Any]= [] 
    for sheet in sheets:
        print(type(wb[sheet]))
        data.append(wb[sheet]) 
    print(data[0]['A3'].value)
    return data

#region Class

class sql_data_handler():
    def __init__(self, config: str) -> None:
        self.config_path = config
        self.config: dict[str,dict[str,str]] = {}
        self.dbconfig: dict[str,str] = {}
        self.cnxn = None
        self.cursor = None
        self.server = ""
        self.database = ""
        self.username = ""
        self.password = ""
        self.driver = "{ODBC Driver 18 for SQL Server}"
        self.trust_server_certificate = "yes"
        self.encrypt = "yes"
        self.timeout = 30
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
    
    def connect(self) -> None:
        self.sql = pyodbc.connect(
            host=self.server,
            user=self.username,
            password=self.password,
            database=self.database,
            driver=self.driver,
            autocommit=True
        )

    def close(self) -> None:
        if self.sql:
            self.sql.close()

#region entry point
if __name__ == "__main__":
    df = read_excel(r"datasheets\20250709 Runsheet.xlsx")
    print(df)

    # temp = sql_data_handler("config.json")
    
    # temp.connect()
    
    # temp.close()

#endregion