import sys
import os

cwd = os.getcwd()

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

exe_path = get_exe_location()
print(f"The executable is located at: {exe_path}, current working path is {cwd}")

if str(cwd).lower() != str(exe_path).lower():
    os.chdir(exe_path)
    print(f"Changed working directory to: {os.getcwd()}")


print(sys.argv)