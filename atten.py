import subprocess
import io
import time
import os, shutil
import pandas as pd
from io import StringIO
from datetime import date
from datetime import timedelta
from PIL import ImageTk,Image
from PIL import ImageGrab
import dataframe_image as dfi
import win32com.client as win32
today = date.today()
yesterday = today - timedelta(days = 1)
today=str(today)
yesterday=str(yesterday)
a='Select EMAIL,ID,MANAGERID,NAME,USERROLEID,FEDERATIONIDENTIFIER FROM User where Country__c'
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"=''IN'' AND ISACTIVE=TRUE AND PROFILEID=''00e3x000001Zh7HAAS'' ' -u -Pro -r csv |out-file atten.csv"
p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
with io.open('atten.csv','r',encoding='utf16') as f:
    text = f.read()
StringData = StringIO(text)
df = pd.read_csv(StringData, sep =",")
df1=df[['Name','Email','Id','UserRoleId','ManagerId','FederationIdentifier']]
df2=df[['Name','Id','Email']]
df2=df2.rename(columns={'Email':'Manager Email'})
df2=df2.drop_duplicates(subset='Id',keep='first')
df2.to_excel('managername.xlsx')
df1=pd.merge(df1,df2,left_on='ManagerId',right_on='Id',how='left')
df1=df1.rename(columns={'Name_x':'Employee Name','Id_x':'User Id','Name_y':'Manager Name'})
usr=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\akhilesh\New user.xlsx")
df1=pd.merge(df1,usr,left_on='UserRoleId',right_on='ID',how='left')
df12=df1[df1['User Type']=='DSM']
df13=df1[df1['User Type']!='DSM']
df13['User Type']=df13['FederationIdentifier'].str.startswith('C')
df13=df13.replace({True:'OffRoll',False:'OnRoll'})
df1=pd.concat([df12,df13],axis=0)
a='Select Attendance_Date_Time__c,CreatedByid,REASON__C,JOINT_WORKING__C,WORKING_WITH_MANUAL__C,WORKING_WITH__C FROM Store_Visit_Related_Info__c WHERE Attendance_Date_Time__c >'+yesterday+'T20:00:01.000Z AND Attendance_Date_Time__c < '+today+'T04:33:59.000Z AND RECORDTYPEID='
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"''0123x000001ZXlfAAG'' ' -u -Pro -r csv |out-file atten.csv"
p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
with io.open('atten.csv','r',encoding='utf16') as f:
    text = f.read()
StringData = StringIO(text)
aten = pd.read_csv(StringData, sep =",")
a='Select Attendance_Date_Time__c,CreatedByid,REASON__C,JOINT_WORKING__C,WORKING_WITH_MANUAL__C,WORKING_WITH__C FROM Store_Visit_Related_Info__c WHERE Attendance_Date_Time__c >'+yesterday+'T20:00:01.000Z AND Attendance_Date_Time__c < '+today+'T04:33:59.000Z AND RECORDTYPEID='
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"''0123x000001ZXleAAG'' ' -u -Pro -r csv |out-file atten1.csv"
p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
with io.open('atten1.csv','r',encoding='utf16') as f:
    text = f.read()
StringData = StringIO(text)
aten1 = pd.read_csv(StringData, sep =",")
aten=pd.concat([aten,aten1])
aten=aten.drop_duplicates(subset='CreatedById',keep='first')
aten.to_excel('atten.xlsx')
df1=pd.merge(df1,aten,left_on='User Id',right_on='CreatedById',how='left')
df1['CreatedById']=df1['CreatedById'].fillna(0)
df1=df1.rename(columns={'Attendance_Date_Time__c':'Attendance_Date_Time','CreatedById':'Status'})
dff=df1[df1['Status']==0]
dfs=df1[df1['Status']!=0]
dff['Status']='Not Marked'
dfs['Status']='Marked'
dfk=pd.concat([dff,dfs],axis=0)
no=pd.read_excel('Noattendence.xlsx')
dfk=pd.merge(dfk,no,on='User Id',how='left')
dfk=dfk.sort_values(by='Employee Name')
dfk['Name']=dfk['Name'].fillna(0)
dfk=dfk[dfk['Name']==0]
dfk=dfk.drop(columns={'Name','ID','NAME','FederationIdentifier','ManagerId','UserRoleId','Id_y'})
dfk=dfk.rename(columns={'Reason__c':'Reason'})
dfk['Region']=dfk['Region'].replace({'India':'South'})
dfk['Biz']=dfk['Biz'].replace({'Full':'B2C'})
Gf=dfk[['Employee Name','User Id']]
Gf=Gf.rename(columns={'User Id':'ID','Employee Name':'Working_With'})
dfk=pd.merge(dfk,Gf,left_on='Working_With__c',right_on='ID',how='left')
dfk=dfk.drop(columns={'Working_With__c','ID'})
dfk=dfk.rename(columns={'Joint_Working__c':'Joint_Working','Working_With_Manual__c':'Working_With_Manual'})
dfk.to_excel('attendence.xlsx',index=False)
dfk=dfk[dfk['Biz']=='B2C']
dfk=dfk.drop(columns={'Biz'})
dfk1=dfk[~dfk['Manager Email'].isin(['milind.acharya@bunge.com','rs.murthy@bunge.com','sanjeev.giri@bunge.com','mahesh.kumar@bunge.com','arun.neogi@bunge.com','sandeep.kaul@bunge.com','rohit.nair@bunge.com','naresh.makhija@bunge.com','bhupendra.singh@bunge.com','swapneswar.dakua@bunge.com','amitkumar.gupta@bunge.com','devinderkumar.sharma@bunge.com','harish.kumarsoni@bunge.com','rajinder.pal@bunge.com'])]
dfk1=dfk1.fillna('NA')
dfk1=dfk1[dfk1['Manager Email']!='NA']
MN=dfk1[['Manager Name','Manager Email']]
MN=MN.drop_duplicates(subset='Manager Email',keep='first')
MN.to_excel('managername.xlsx')
dfk2=dfk[dfk['Manager Email'].isin(['bhupendra.singh@bunge.com','swapneswar.dakua@bunge.com','amitkumar.gupta@bunge.com','devinderkumar.sharma@bunge.com','harish.kumarsoni@bunge.com','rajinder.pal@bunge.com'])]
dfk2=dfk2.fillna('NA')
dfk2=dfk2[dfk2['Manager Email']!='NA']
MN1=dfk2[['Manager Name','Manager Email']]
MN1=MN1.drop_duplicates(subset='Manager Email',keep='first')
MN1.to_excel('managername1.xlsx')
dfk=dfk[['Employee Name','Email','Manager Name','Manager Email','Region','User Type','Status','Reason','Working_With']]
for ind in MN1.index:
    a=MN1['Manager Email'][ind]
    b=MN1['Manager Name'][ind]
    dfk1=dfk[dfk['Manager Email']==a]
    dfk1=dfk1.drop(columns={'Manager Email','Manager Name','Email'})
    dfk1=dfk1.sort_values(["User Type"], ascending = True)
    dfi.export(dfk1, 'dataframe.png')
    dataframe=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\akhilesh\dataframe.png"
    html_body=r"""Dear """ +b+""",
    <p>Please check the attendance status and reason for all of your users listed below. Attendance has been considered before 10 AM .</p>
    <H3><u>Attendance</u></H3>
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
    mail.Subject = "SFA attendance report | "+today  
    mail.SentOnBehalfOfName='bas.in.salestech.support@bunge.com'
    mail.Cc= 'rs.murthy@bunge.com;'
    doc=inspector.WordEditor
    selection=doc.Content
    selection.Find.Text=r"{Image1}"
    selection.Find.Execute()
    selection.Text=""
    selection.Text
    img=selection.InlineShapes.AddPicture(dataframe,0,1)
    per = 40
    img.Height = int(per*35.581)
    img.Width  = int(per*11.100)
    #mail.Send()
    time.sleep(5)
for ind in MN.index:
    a=MN['Manager Email'][ind]
    b=MN['Manager Name'][ind]
    dfk1=dfk[dfk['Manager Email']==a]
    dfk1=dfk1.drop(columns={'Manager Email','Manager Name','Email'})
    dfk1=dfk1.sort_values(["User Type"], ascending = True)
    dfi.export(dfk1, 'dataframe.png')
    dataframe=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\akhilesh\dataframe.png"
    html_body=r"""Dear """ +b+""",
    <p>Please check the attendance status and reason for all of your users listed below. Attendance has been considered before 10 AM .</p>
    <H3><u>Attendance</u></H3>
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
    mail.Subject = "SFA attendance report | "+today  
    mail.SentOnBehalfOfName='bas.in.salestech.support@bunge.com'
    mail.Cc= 'rs.murthy@bunge.com;'
    doc=inspector.WordEditor
    selection=doc.Content
    selection.Find.Text=r"{Image1}"
    selection.Find.Execute()
    selection.Text=""
    selection.Text
    img=selection.InlineShapes.AddPicture(dataframe,0,1)
    #mail.Send()
    time.sleep(5)



