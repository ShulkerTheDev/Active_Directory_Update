import pandas as pd
import subprocess
import re

#Pass excel sheet into a dataframe
df = pd.read_excel('Excel_Sheets//Exec List at Jan 2023.xlsx', keep_default_na=False)
df2 = pd.read_excel('Excel_Sheets//Manager list at Jan 2023.xlsx', keep_default_na=False)
df3 = pd.read_excel('Excel_Sheets//Non-Manager List at Jan 2023.xlsx', keep_default_na=False)

#create file to store error logs
logFile = open('errors.txt', 'w') 

#Loops over each row in the dataframe
for index, row in df2.iterrows():
  employeeName = row["Name"] 

  employeeFName = re.split(', | ', employeeName)[1]
  employeeLName = re.split(', | ', employeeName)[0]

  #Check if middle inital exists if so assign it, if not assign None
  try:
    employeeMInital = re.split(', | ', employeeName)[2]
  except IndexError:
    employeeMInital = None

  try:
    command = f"Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; Remove-DistributionGroupMember -Identity 'UPO Members' -Member '{employeeFName} {employeeLName}'"
    proc =subprocess.check_output(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", command], shell=True, input= "y", text=True)
  except Exception as e:
    if(employeeMInital is not None):
      try:
        command = f"Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; Remove-DistributionGroupMember -Identity 'UPO Members' -Member '{employeeFName} {employeeMInital}. {employeeLName}'"
        proc =subprocess.check_output(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", command], shell=True, input= "y", text=True)
      except Exception as e:
        logFile.write(str(e)+"\n")
    else:
      logFile.write(str(e)+"\n")



