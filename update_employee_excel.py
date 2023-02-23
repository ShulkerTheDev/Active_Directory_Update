import pandas as pd
import subprocess
from getpass import getpass
import re

#Pass excel sheet into a dataframe
df = pd.read_excel('TCMD LIst.xls', keep_default_na=False)
#password = getpass()

#Loops over each row in the dataframe
for index, row in df.iterrows():
    employeeLName = row["Last Name"] 
    employeeUsrName = row["Network Username"]
    employeeEmail = row["mail"]
    employeeDepartment = row["department"]
    #print(employeeUsrName)

    try:
        proc =subprocess.check_output(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", f"Get-AdUser '{employeeUsrName}' -Properties Department, Mail, Surname | fl Surname, SamAccountName, Mail, Department"]).decode("utf-8")
    
        employeeSurname = re.split(':|\n', proc)[3]
        employeeMail = re.split(':|\n', proc)[7]
        employeeDepart = re.split(':|\n', proc)[9]
        
        print(row)
        df.at[index, "Last Name"]=employeeSurname
        df.at[index, "mail"]=employeeMail
        df.at[index, "department"]=employeeDepart
    except Exception as e:
        message = str(e)


df.to_excel("output.xlsx")

