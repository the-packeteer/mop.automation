import os
from pathlib import Path
import pandas as pd

# Function that using dictionary "replacements" to replace vars in Template
def replace_all(text, dic):
    for i, j in dic.items():
        text = text.replace(i, j)
    return text


# Function to convert lists to strings
def listToString(s):
    # initialize an empty string
    str1 = " "
    # return string
    return str1.join(s)


# READING EXCEL SPREADSET
# Destination Folder for Completed MOP
PID = os.environ['USERPROFILE']
# Locations for Needed Files
mopCreator = Path(PID + '\\Documents\\mop.generator\\mop.generator.xlsm')
outputLocation = Path(PID + '\\Documents\\mop.generator\\completed.mops')
outputTemplate = Path(PID + '\\Documents\mop.generator\\output.html')
# Open Excel spreadsheet
mopCreatorRead = pd.read_excel((mopCreator), engine='openpyxl')
# Reads Template Values Column in Spreadsheet and adds to list 'varList' and 'valueList'
varList = mopCreatorRead['Template Variables'].tolist()
valueList = mopCreatorRead['Template Values'].tolist()
templatePath = mopCreatorRead['Template Path'].tolist()  # adds Template Path into 'templatePath'
# Removes cells with no data
VarsList = [x for x in varList if str(x) != 'nan']
valueList = [x for x in valueList if str(x) != 'nan']
existList = [x for x in templatePath if str(x) != 'nan']
# Turns all entries in 'varList' and 'valueList' into strings
VarsList = [str(i) for i in VarsList]
valueList = [str(i) for i in valueList]
# Removes heading and trailing spaces in 'varList' and 'valueList'
VarsList = [x.strip() for x in VarsList]
valueList = [x.strip() for x in valueList]
# combines 'mopVarsList1' and 'mopVarsList2' into a Dictionary, called 'mopVarDict'
replacements = dict(zip(VarsList, valueList))
# Converts 'existList' into string, strips quotation marks, and adds to var named 'template_Path'
templatePathName = Path(str(templatePath[0]).strip('"'))
templatePathNameCheck = str(templatePathName)
list = templatePathNameCheck.split('\\Documents')
notNeeded, pathVAR = list
pathVar = {'{PATH_VAR}': pathVAR}
# Check if Existing path is provided
with open(templatePathName, 'r', encoding='utf8') as infile, open(outputTemplate, 'w', encoding='utf8') as outfile:
    # reads var 'infile' to var 'boilFileConversion'
    boilFileConversion = infile.read()
    # uses 'replace_all' function to replace 'Template Variables' with 'Template Values' in var 's', and store in
    # new var 'replace'
    replace = replace_all(boilFileConversion, replacements)
    outfile.write(replace)
