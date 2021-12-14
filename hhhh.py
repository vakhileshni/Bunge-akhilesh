import subprocess
import io
import pandas as pd
from io import StringIO
from datetime import date
from datetime import timedelta
today = date.today()
yesterday = today - timedelta(days = 1)
yesterday=str(yesterday)
#a='Select COUNTRY__C,EMAIL,ID,ISACTIVE,MANAGERID,NAME,PROFILEID,USERROLEID,FEDERATIONIDENTIFIER FROM User'
#a=str(a)
#b="sfdx force:data:soql:query -q '"+a+"' -u -Pro -r csv"
#p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
with io.open('atten.csv','r',encoding='utf16') as f:
    text = f.read()
StringData = StringIO(text)
User= pd.read_csv(StringData, sep =",")
df=User[['COUNTRY__C','EMAIL','ID','ISACTIVE','MANAGERID','NAME','PROFILEID','USERROLEID','FEDERATIONIDENTIFIER']]
df=df[df['COUNTRY__C']=='IN']
df=df.replace({True:1,False:0})
df=df[df['ISACTIVE']==1]
df=df[df['PROFILEID']=='00e3x000001Zh7HAAS']
df1=df[['NAME','EMAIL','ID','USERROLEID','MANAGERID','FEDERATIONIDENTIFIER']]
df2=df[['NAME','MANAGERID']]
df2=df2.drop_duplicates(subset='MANAGERID',keep='first')
df1=pd.merge(df1,df2,on='MANAGERID',how='left')
df1=df1.rename(columns={'NAME_x':'EMPLOYEE NAME','ID':'USER ID','NAME_y':'MAMAGER NAME'})
usr=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\akhilesh\New user.xlsx")
df1=pd.merge(df1,usr,left_on='USERROLEID',right_on='ID',how='left')
df1['User Type']=df1['FEDERATIONIDENTIFIER'].str.startswith('U')
df1=df1.replace({True:'OnRoll',False:'OffRoll'})
df1.to_excel('attendence.xlsx',index=False)
