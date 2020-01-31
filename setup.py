import sys
from cx_Freeze import setup, Executable

base = 'Win32GUI'

build_exe_options = {'packages': ['os','tkinter','openpyxl','datetime'], 'include_files': ['scu.ico']}

exec = Executable('process_gui.py',
                  targetName = 'excel_processor.exe',
                  base = base,
                  icon = 'scu.ico')

setup (name = 'Quarterly Excel Processor',
       version = '1.0',
       description = 'Process Full Quarters in Excel for the One Stop Office',
       options = {'build_exe':build_exe_options},
       executables = [exec])
