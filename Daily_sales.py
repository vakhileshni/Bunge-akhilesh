import subprocess
import io
import time
import pandas as pd
import numpy as np
from io import StringIO
from datetime import date
from datetime import timedelta
import dataframe_image as dfi
from PIL import ImageTk,Image
from PIL import ImageGrab
import win32com.client as win32
today = date.today()
today1=str(today)
yesterday = today - timedelta(days = 1)
yesterday=str(yesterday)
master=pd.read_excel('Master.xlsx')
a='Select EMAIL,ID,MANAGERID,NAME,USERROLEID FROM User where Country__c'
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"=''IN'' AND ISACTIVE=TRUE AND PROFILEID=''00e3x000001Zh7HAAS'' ' -u -Pro -r csv |out-file atten.csv"
p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
with io.open('atten.csv','r',encoding='utf16') as f:
    text = f.read()
StringData = StringIO(text)
df = pd.read_csv(StringData, sep =",")
df1=df[['Name','Email','Id','UserRoleId','ManagerId']]
df2=df[['Name','Id','Email']]
df2=df2.rename(columns={'Email':'Manager Email'})
df2=df2.drop_duplicates(subset='Id',keep='first')
df1=pd.merge(df1,df2,left_on='ManagerId',right_on='Id',how='left')
df1=df1.rename(columns={'Name_x':'Employee Name','Id_x':'User Id','Name_y':'Manager Name'})
usr=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\akhilesh\New user.xlsx")
df1=pd.merge(df1,usr,left_on='UserRoleId',right_on='ID',how='left')
U_Master=df1[['User Id','Email','Manager Name','Manager Email','Region','User Type','Biz']]
master=master.drop_duplicates(subset=['Material'],keep='first')
agent=pd.read_excel('sales_order.xlsx')
agent=pd.merge(agent,master,left_on='Product: SAP Number',right_on='Material',how='left')
agent=agent[['Check In','Check Out','Owner Business Line','Last Modified By: Full Name','Last Modified By: Case-Safe User ID','Store Visit Name','Store Visit Order Name','Store Visit Order Product Name','Status','No Order Reason','Product: SAP Number','Product: Product Name','Remote Visit','Account: Account Name','Account: City: Store Visit Related Info Name','Owner Business Zone','Account: State','Account: Region','Distributor: SAP Account Number','Distributor: Account Name','Beat','Order Quantity (Case)','Order Quantity (Piece)','Gross Weight','Pack size']]
agent=agent.fillna(0)
agent['Check']=agent['Check In'].astype('str')+agent['Check Out'].astype('str')
agent=agent[agent['Check']!='00']
agent=agent.drop(columns='Check')
agent=pd.merge(agent,U_Master,left_on='Last Modified By: Case-Safe User ID',right_on='User Id',how='left')
agent['RBM']=agent['Region'].replace({'East':'Arun Neogi','Chambal':'Mahesh Kumar','North 1':'Sandeep Kaul','North 2':'Rohit Nair','South':'RS Murthy','West':'Naresh Makhija'})
agent=agent.sort_values(['Last Modified By: Full Name','Store Visit Order Name'])
agent=agent.rename(columns={'Last Modified By: Full Name':'Store Visitor Name','Product: SAP Number':'Product SAP Number','Product: Product Name':'Product Name','Account: Account Name':'Retailer Name','Account: City: Store Visit Related Info Name':'Retailer City','Account: State':'Retailer State','Account: Region':'Retailer Region','Distributor: SAP Account Number':'Distributor SAP Number','Distributor: Account Name':'Distributor Name'})
agent=agent[['Store Visitor Name','Email','Manager Name','Manager Email','Region','User Type','RBM','Store Visit Name','Store Visit Order Name']]
ASM=agent[['Manager Name','Manager Email']]
ASM=ASM.drop_duplicates(subset='Manager Email',keep='first')
ASM.to_excel('ASM.xlsx',index=False)
SM=agent[['Store Visitor Name','Email','Manager Email','Region','User Type']]
SM=SM.drop_duplicates(subset='Email',keep='first')
SVO=agent[['Email','Store Visit Name']]
SVO=SVO.drop_duplicates(subset='Store Visit Name',keep='first')
SVO=SVO.groupby(['Email']).count()
SVO=SVO.reset_index()
SM=pd.merge(SM,SVO,on='Email',how='left')
PC=agent[['Email','Store Visit Order Name']]
PC=PC.drop_duplicates(subset='Store Visit Order Name',keep='first')
PC=PC[PC['Store Visit Order Name']!=0]
PC=PC.groupby(['Email']).count()
PC=PC.reset_index()
SM=pd.merge(SM,PC,on='Email',how='left')
SM=SM.rename(columns={'Store Visit Name':'TC','Store Visit Order Name':'PC'})
SM['PC']=SM['PC'].fillna(0)
SM['PC']=SM['PC'].round(0)
DS=pd.read_excel('Distributor.xlsx')
DS=DS[['Email','Distributor Name']]
SM=pd.merge(SM,DS,on='Email',how='left')
SM2=pd.merge(SM,ASM,on='Manager Email',how='left')
SM2=SM2[SM2['User Type']=='DSM']
SM2=SM2[['Store Visitor Name','Email','Manager Name','Manager Email','Region','User Type','TC','PC','Distributor Name']]
Atten=pd.read_excel('attendence.xlsx')
Atten=Atten[['Employee Name','Email','Manager Name','Manager Email','Region','User Type']]
Atten=Atten.drop_duplicates(subset='Email',keep='first')
#Atten[TC]=0
#Atten[PC]=0
Atten=Atten.rename(columns={'Employee Name':'Store Visitor Name'})
Atten=Atten[Atten['User Type']=='DSM']
SM3=SM2[['Email']]
SM3=SM3.rename(columns={'Email':'Email1'})
Atten=pd.merge(Atten,SM3,left_on='Email',right_on='Email1',how='left')
Atten=Atten.fillna(0)
Atten.to_excel('akhiles.xlsx',index=False)
Atten=Atten[Atten['Email1']==0]
Atten=Atten.drop(columns={'Email1'})
SM2=pd.concat([SM2,Atten],axis=0)
SM2=SM2.drop(columns={'Distributor Name'})
SM2=pd.merge(SM2,DS,on='Email',how='left')
SM2.to_excel('Sales Productivity.xlsx',index=False)
for ind in ASM.index:
    a=ASM['Manager Email'][ind]
    b=ASM['Manager Name'][ind]
    SM1=SM[SM['Manager Email']==a]
    SM1=SM1.drop(columns={'Manager Email','Email'})
    SM1=SM1.sort_values(["User Type"], ascending = True)
    dfi.export(SM1, 'SMframe.png')
    dataframe=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\akhilesh\SMframe.png"
    html_body=r"""Dear """ +b+""",
    <p>Please check the sales productivity status for all of your users listed below.</p>
    <H3><u>Sales Productivity</u></H3>
    {Image1}
    <br></br>
    <br></br>
    <br></br>
    <br></br>
    <H4> Thanks & Regards,</H4>
    <H5> Akhilesh Pal </H5>
"""
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.HTMLBody=html_body
    inspector=mail.GetInspector
    inspector.Display()
    mail.To =a
    mail.Subject ="Sales productivity status | " +today1 
    mail.SentOnBehalfOfName='bas.in.salestech.support@bunge.com'
    mail.Cc= 'rs.murthy@bunge.com;'
    doc=inspector.WordEditor
    selection=doc.Content
    selection.Find.Text=r"{Image1}"
    selection.Find.Execute()
    selection.Text=""
    selection.Text
    img=selection.InlineShapes.AddPicture(dataframe,0,1)
    mail.Send()
    time.sleep(5)