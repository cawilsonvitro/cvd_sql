import winreg
import os 
import json
import sys

def get_exe_location():
    """
    Returns the absolute path to the compiled executable.
    """
    if getattr(sys, 'frozen', False):
        # Running in a PyInstaller bundle
        return os.path.dirname(sys.executable)
    else:
        # Running as a regular Python script
        return os.path.dirname(os.path.abspath(__file__))

# exe_loc = get_exe_location().split(os.sep)

# exe_loc = os.sep.join(exe_loc) 
# main_exe = "excel2sql.exe"
# print("exe location:", exe_loc)
# #setting up env
# key_path = "Environment"
# root_key = winreg.HKEY_CURRENT_USER
# try:
#     variable_name = "Path"
#     key = winreg.OpenKey(root_key, key_path, 0, winreg.KEY_SET_VALUE | winreg.KEY_READ | winreg.KEY_WRITE)
#     current_path, reg_type = winreg.QueryValueEx(key, variable_name)
#     new_path_to_add = exe_loc + "\\"
    
#     if new_path_to_add not in current_path:
#         updated_path = current_path + ";" + new_path_to_add
#         print(updated_path)
#         winreg.SetValueEx(key, variable_name, 0, reg_type, updated_path)
#     # key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ | winreg.KEY_WRITE)
# except OSError as e:
#     key = ""
#     print(f"Error opening registry key: {e}")
#     # Handle cases where administrative privileges are required for system variables

# print(key)



# variable_value = exe_loc + "\\" + main_exe


# os.environ["Path"] = "test"





# setting up json
with open('config_def.json', 'r') as f:
    config = json.load(f)
    
config['Database_Config']['host'] = input("Enter your host (default is localhost): ")
config['Database_Config']['db'] = input("Enter your database name (default is cvd_test): ")
config['Database_Config']['username'] = input("Enter your username for the database: ")
config['Database_Config']['password'] = input("Enter your password for the database: ")

# with open('config.json', 'w') as f:
#     json.dump(config, f, indent=4)

