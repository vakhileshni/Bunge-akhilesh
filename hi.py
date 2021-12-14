import subprocess
import io
import pandas as pd
from io import StringIO
from datetime import date
from datetime import timedelta
today = date.today()
yesterday = today - timedelta(days = 1)
yesterday=str(yesterday)
a='Select EMAIL,ID,MANAGERID,NAME,USERROLEID,FEDERATIONIDENTIFIER FROM User where Country__c'
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"=''IN'' AND ISACTIVE=True AND PROFILEID=''00e3x000001Zh7HAAS'' ' -u -Pro -r csv |out-file atten.csv"
subprocess.run(['powershell',b],shell=True)