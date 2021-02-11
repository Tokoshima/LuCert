import subprocess
import sys
#subprocess.check_call([sys.executable, "-m", "pip", "install", "manim"])

subprocess.check_call([sys.executable, "-m", "pip", "install","openpyxl"])
subprocess.check_call([sys.executable, "-m", "pip", "install","python-docx"])
subprocess.check_call([sys.executable, "-m", "pip", "install", "docx"])

subprocess.check_call([sys.executable, "-m", "pip", "install", "PyQt5"])
# def install(package):
#     subprocess.check_call([sys.executable, "-m", "pip3", "install", "manim",package])
#     subprocess.check_call([sys.executable, "-m", "pip", "install","openpyxl", package])
#     subprocess.check_call([sys.executable, "-m", "pip", "install","python-docx", package])
#     subprocess.check_call([sys.executable, "-m", "pip", "install", "docx",package])
#
#     subprocess.check_call([sys.executable, "python3", "-m", "pip", "install", "PyQt5", package])
