#region imports
import os
# import typing
import venv #type:ignore



#endregion


def venv_builder(req:str = "constraints.txt") -> None:
    """
    Creates a Python virtual environment in the current working directory and installs dependencies from a requirements file.
    This function performs the following steps:
    1. Reads the specified requirements file (default: "constraints.txt").
    2. Strips whitespace from each line and replaces any line containing "delcom" with a local wheel file path.
    3. Writes the processed requirements back to the file.
    4. Creates a new virtual environment in a ".venv" directory if it does not already exist.
    5. Installs the dependencies from the requirements file into the virtual environment using pip.
    Args:
        req (str, optional): Path to the requirements file. Defaults to "constraints.txt".
    Returns:
        None
    """
    cwd = os.getcwd()
    venv_path = os.path.join(cwd, '.venv') 
    if not os.path.exists(venv_path):

        venv.create(venv_path, with_pip=True, clear=True)
        
        script = os.path.join(venv_path, 'Scripts')

        py = os.path.join(script, 'python.exe')

        pip = os.path.join(script, 'pip.exe')
        

        install = f"{py} {pip} install -r requirements.txt"

        os.system(install)

        
        
        

if __name__ == "__main__":
    try:
        venv_builder()
    except Exception as e:
        print("Exception in main: " + str(e))