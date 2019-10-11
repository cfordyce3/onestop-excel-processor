import sys
from cx_Freeze import setup, Executable

build_exe_options = {'build_exe':'..\\process_month\\','packages': ['os','tkinter','openpyxl'], 'include_files': ['scu.ico']}

base = None
if (sys.platform == 'win32'):
    base = 'Win32GUI'

exec = Executable('process_month_gui_standard.py',
                  targetName = 'process_month.exe',
                  base = base,
                  icon = 'scu.ico')

setup (name = 'Process Month',
       version = '1.0',
       description = 'Process month application for One Stop',
       options = {'build_exe':build_exe_options},
       executables = [exec])
