@echo off
echo Installing necessary components for the excel processor...

\dependencies\python-3.7.4.exe /quiet InstallAllUsers=1 TargetDir=%ProgramFiles%\Python3.7.4 Include_pip=1 Include_test=0 PrependPath=1
%ProgramFiles%\Python3.7.4\python.exe -m pip install openpyxl

pause