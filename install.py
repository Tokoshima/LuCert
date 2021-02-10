import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install","openpyxl", package])
    subprocess.check_call([sys.executable, "-m", "pip", "install","python-docx", package])
    subprocess.check_call([sys.executable, "-m", "pip", "install", "docx",package])
    subprocess.check_call([sys.executable, "python3", "-m", "pip", "install", "PyQt5", package])
