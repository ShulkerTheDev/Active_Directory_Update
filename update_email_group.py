import pandas as pd
import subprocess
import re
import getpass

#Pass excel sheet into a dataframe
df = pd.read_excel('Excel_Sheets//Exec List at Jan 2023.xlsx', keep_default_na=False)
df2 = pd.read_excel('Excel_Sheets//Manager list at Jan 2023.xlsx', keep_default_na=False)
df3 = pd.read_excel('Excel_Sheets//Non-Manager List at Jan 2023.xlsx', keep_default_na=False)

#Prompts for password
#password = getpass.getpass('Password:')

#Loops over each row in the dataframe
for index, row in df.iterrows():
    employeeName = row["Name"] 

    employeeFName = re.split(', | ', employeeName)[1]
    employeeLName = re.split(', | ', employeeName)[0]

    try:
        proc =subprocess.check_output(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", "Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn", f"Add-DistributionGroupMember -Identity 'Management Staff' -Member '{employeeFName} {employeeLName}'"]).decode("utf-8")
    except Exception as e:
        print(str(e))


