#region imports
import json
import socket
import os
# import typing
import sys
import subprocess
import venv #type:ignore
import traceback
import logging
from logging.handlers import TimedRotatingFileHandler
import datetime as dt


#endregion


date = dt.now().strftime("%m-%d-%Y, Hour %H Min %M Sec %S")
logging.basicConfig(
    level=logging.DEBUG, # Set a global logging level
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        # logging.StreamHandler(), # Log to console
        TimedRotatingFileHandler(f'logs\\{date}.log', when = "D", backupCount= 5)
    ]
)


def venv_builder(req = "constraints.txt") -> None:
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
    global logger
    logger.debug("Starting virtual environment builder")
    logger.debug(f"Using requirements file: {req}")
    lines: list[str]
    req_file:str = req
    with open( req_file, 'r') as f:
        lines = list(f.readlines())
    logger.debug(f"Read {len(lines)} lines from requirements file")

    logger.debug("Stripping lines and replacing 'delcom' entries")
    stripped_lines: list[str] = []
    stripped: str = ""
    cwd = os.getcwd()
    venv_path = os.path.join(cwd, '.venv') 
    if not os.path.exists(venv_path):
        for line in lines:
            stripped = line.strip()
            if "delcom" in stripped:
                stripped = "delcom @ file:///" + os.path.join(cwd,"install_files","delcom-0.1.1-py3-none-any.whl")
            if stripped != "":
                stripped_lines.append(stripped)

        with open(req_file, "w") as f:
            for line in stripped_lines:
                f.write("\n" + line)
                
        logger.debug(f"Creating virtual environment at {venv_path}")
        venv.create(venv_path, with_pip=True, clear=True)

        logger.debug("Virtual environment created, installing dependencies")
        
        script = os.path.join(venv_path, 'Scripts')

        py = os.path.join(script, 'python.exe')

        pip = os.path.join(script, 'pip.exe')
        
        logger.debug(f"Using Python interpreter at {py} and pip at {pip}")
        install = f"{py} {pip} install -r constraints.txt"

        os.system(install)
        logger.debug("Dependencies installed")