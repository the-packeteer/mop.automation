import os
from pathlib import Path
import pandas as pd
import subprocess

# COPY/PASTE FOR EACH SCRIPT TAG (i.e <vbs></vbs> or <python></python>
# Grabs current user userID
PID = os.environ['USERPROFILE']
# Locations for Needed Files
outputTemplate = Path(PID + '\\Documents\mop.generator\\output.html')
destinationFile = Path(PID + '\\Documents\\\SecureCRT\\Scripts\\Pre-Check_Output.vbs')
# Reads 'completed.mop.html' and parses out <vbs></vbs> tag pairs into 'Pre-Check_Output.vbs'
with open(outputTemplate, 'r', encoding='utf8') as infile, open(destinationFile, 'w', encoding='utf8') as outfile:
    # Parses varable 'infile', and copies data between <></> tags
    copy = False
    for line in infile:
        if line.strip() == "<vbs>":
            copy = True
            continue
        elif line.strip() == "</vbs>":
            copy = False
            continue
        elif copy:
            outfile.write(line)  # writes data to file specified in 'destinationFile' variable