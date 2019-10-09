import sys
from cx_Freeze import setup, Executable

build_exe_options = {'packages': ['os','tkinter','openpyxl']}

base = None
if (sys.platform == 'win32'):
    base = 'Win32GUI'

setup (name = 'Process Month', version = '1', description = 'Process month application for One Stop', options = {'build_exe':build_exe_options}, executables = [Executable('process_month_gui.py', base=base)])
