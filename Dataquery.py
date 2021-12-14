import subprocess
import io
from datetime import date
from datetime import timedelta
today = date.today()
yesterday = today - timedelta(days = 1)
yesterday=str(yesterday)
a='Select CreatedById,LastModifiedById,Name FROM Store_Visit__c WHERE LASTMODIFIEDDATE > '+yesterday+'T00:00:01.000Z'
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"' -u -Pro -r csv"
from io import StringIO
import pandas as pd
p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
#with io.open('SVO.csv','r',encoding='utf16') as f:
#    text = f.read()
StringData = StringIO(p.stdout)
df = pd.read_csv(StringData, sep =",")
df.to_excel('SVO.xlsx',index=False)