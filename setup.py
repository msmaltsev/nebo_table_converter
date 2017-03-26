import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os", "openpyxl", "xlrd", "xlwt"], 'include_files':['source/',
                'results/',
                'xml_results/', 'contragent.xlsx']}

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(  name = "excel_process",
        version = "0.1",
        description = "excel_process",
        options = {"build_exe": build_exe_options},
        executables = [Executable("main.py", base=base)])

