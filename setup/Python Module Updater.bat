REM this batch file will download the needed Python modules to run the automated workflow
@ECHO ON
python -m pip install --upgrade pip --user
python -m pip install pandas --user
python -m pip install openpyxl --user
python -m pip install pathlab --user
python -m pip install xlrd --user
ECHO Your Python environment has been updated successfully. 
PAUSE


