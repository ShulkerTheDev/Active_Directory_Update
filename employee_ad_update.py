import pandas as pd
import subprocess
import re

#Pass excel sheet into a dataframe
df = pd.read_excel('Employee List for August 2023 final for IT.xlsx', keep_default_na=False)

#Loops over each row in the dataframe
for index, row in df.iterrows(): 
    #print (row["Name"]) 
    employeeName = row["Name"]
    try:
      employeeFName = re.split(', | ', employeeName)[0]
    except IndexError:
      continue
    try :
      employeeLName = re.split(', | ', employeeName)[1]
    except IndexError:
      continue
    employeeDescription = row["Position Title"]
    employeeDepartment = row["Department"]
    employeeDivision = row["Division"]


    #subprocess.call(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", f"Get-ADUser -Filter 'Name -like ''{employeeFName}*{employeeLName}''' -Properties Division, Department, Description| Fl *"])
    subprocess.run(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", "RunAs /user:nib-bahamas\z-devaughn",f"Get-ADUser -Filter 'Name -like ''{employeeFName}*{employeeLName}''' | Set-ADUser -Description '{employeeDescription}' -Department '{employeeDepartment}' -Division '{employeeDivision}'"])
    #subprocess.call(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", f"Get-ADUser -Filter 'Name -like ''{employeeFName}*{employeeLName}''' -Properties Division, Department, Description| Fl *"])
