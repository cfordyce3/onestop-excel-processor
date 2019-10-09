@echo off
echo Installing necessary components for the excel processor...
echo Installing Python 3.7.4

cd dependencies
python-3.7.4.exe InstallAllUsers=1 /quiet TargetDir="C:\Program Files\Python3.7.4" Include_pip=1 Include_test=0 PrependPath=1
echo Done
echo Installing necessary libraries
cd "C:\Program Files\Python3.7.4"
python.exe -m pip install --upgrade pip --user
python.exe -m pip install openpyxl --user
echo Done

pause