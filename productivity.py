import subprocess
import io
import pandas as pd
from io import StringIO
from datetime import date
from datetime import timedelta
import warnings
warnings.filterwarnings('ignore')
today = date.today()
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
df2=df[['Name','Id']]
df2=df2.drop_duplicates(subset='Id',keep='first')
df2.to_excel('managername.xlsx')
df1=pd.merge(df1,df2,left_on='ManagerId',right_on='Id',how='left')
df1=df1.rename(columns={'Name_x':'Employee Name','Id_x':'User Id','Name_y':'Manager Name'})
usr=pd.read_excel("New user.xlsx")
df1=pd.merge(df1,usr,left_on='UserRoleId',right_on='ID',how='left')
U_Master=df1[['User Id','Employee Name','Email','Manager Name','Region','User Type','Biz']]
U_Master=U_Master[U_Master['Biz']=='B2C']
U_Master=U_Master[U_Master['User Type'].isin(['DSM','Field Force'])]
master=master.drop_duplicates(subset=['Material'],keep='first')
agent=pd.read_excel('B2C_agent.xlsx')
agent=pd.merge(agent,master,left_on='Product: SAP Number',right_on='Material',how='left')
agent=agent[['Check In','Check Out','Owner Business Line','Last Modified By: Full Name','Last Modified By: Case-Safe User ID','Store Visit Name','Store Visit Order Name','Store Visit Order Product Name','Status','No Order Reason','Product: SAP Number','Product: Product Name','Remote Visit','Account: Account Name','Account: City: Store Visit Related Info Name','Owner Business Zone','Account: State','Account: Region','Distributor: SAP Account Number','Distributor: Account Name','Beat','Order Quantity (Case)','Order Quantity (Piece)','Gross Weight','Pack size']]
agent=agent.fillna(0)
agent['Check']=agent['Check In'].astype('str')+agent['Check Out'].astype('str')
agent=agent[agent['Check']!='00']
agent=agent.drop(columns='Check')
agent1=agent[agent['Order Quantity (Piece)']>0]
agent2=agent[agent['Order Quantity (Piece)']==0]
agent1['case']=agent1['Order Quantity (Piece)']/agent1['Pack size']
agent1=agent1.drop(['Order Quantity (Case)','Order Quantity (Piece)'],axis=1)
agent1=agent1.rename(columns={'case':'Order Quantity (Case)'})
agent2=agent2.drop(['Order Quantity (Piece)'],axis=1)
agent3=pd.concat([agent1,agent2],axis=0)
agent3['Total Liter']=agent3['Gross Weight']*agent3['Order Quantity (Case)']
agent3=pd.merge(agent3,U_Master,left_on='Last Modified By: Case-Safe User ID',right_on='User Id',how='left')
agent3['RBM']=agent3['Region'].replace({'East':'Arun Neogi','Chambal':'Mahesh Kumar','North 1':'Sandeep Kaul','North 2':'Rohit Nair','South':'RS Murthy','West':'Naresh Makhija'})
agent3=agent3.sort_values(['Last Modified By: Full Name','Store Visit Order Name'])
agent3=agent3.rename(columns={'Last Modified By: Full Name':'Store Visitor Name','Product: SAP Number':'Product SAP Number','Product: Product Name':'Product Name','Account: Account Name':'Retailer Name','Account: City: Store Visit Related Info Name':'Retailer City','Account: State':'Retailer State','Account: Region':'Retailer Region','Distributor: SAP Account Number':'Distributor SAP Number','Distributor: Account Name':'Distributor Name'})
agent3=agent3[['Check In','Check Out','Owner Business Line','Store Visitor Name','Email','Manager Name','Region','User Type','RBM','Store Visit Name','Store Visit Order Name','Store Visit Order Product Name','Status','No Order Reason','Product SAP Number','Product Name','Remote Visit','Retailer Name','Retailer City','Owner Business Zone','Retailer State','Retailer Region','Distributor SAP Number','Distributor Name','Beat','Gross Weight','Pack size','Order Quantity (Case)','Total Liter']]
agent3.to_excel('daily.xlsx',index=False)
nw=agent3[['Email','Region','User Type']]
nw=nw.drop_duplicates(subset=['Email'],keep='first')
nw1=nw[['Email','Region']]
nw1=nw1.rename(columns={'Region':'check'})
U_Master0=pd.merge(U_Master,nw1,on='Email',how='left')
U_Master0=U_Master0.fillna(0)
U_Master0=U_Master0[U_Master0['check']==0]
U_Master0=U_Master0.drop(['check'],axis=1)
U_Master0.to_excel('Not_working_poeple.xlsx',index=False)
Format=pd.read_excel('Format.xlsx')
Format['Region-User Type']=Format['Region']+Format['User Type']
agent3['Region-User Type']=agent3['Region']+agent3['User Type']
U_Master['Region-User Type']=U_Master['Region']+U_Master['User Type']
U_Master1=U_Master[U_Master['Biz']=='B2C']
U_Master1.to_excel('FFDSM.xlsx.xlsx',index=False)
U_Master1=U_Master1[U_Master1['User Type'].isin(['Field Force','DSM'])]
U_Master2=U_Master1[['User Id','Region-User Type']]
U_Master2=U_Master2.groupby(['Region-User Type']).count()
U_Master2=U_Master2.reset_index()
Format=pd.merge(Format,U_Master2,on='Region-User Type',how='left')
nw['Region-User Type']=nw['Region']+nw['User Type']
nw2=nw[['Region-User Type','Email']]
nw2=nw2.groupby(['Region-User Type']).count()
nw2=nw2.reset_index()
nw2=nw2.rename(columns={'Email':'Working People'})
Format=pd.merge(Format,nw2,on='Region-User Type',how='left')
Format['TC Target']=Format['User Id']*40
agent4=agent3[['Region-User Type','Total Liter']]
agent4['Region-User Type']=agent4['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
agent4=agent4.groupby(['Region-User Type']).sum()
agent4=agent4.reset_index()
Format=pd.merge(Format,agent4,on='Region-User Type',how='left')
Format=Format.rename(columns={'Total Liter':'Liter Sold'})
agent5=agent3[['Region-User Type','Store Visit Name']]
agent5['Region-User Type']=agent5['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
agent5=agent5.drop_duplicates(subset=['Store Visit Name'],keep='first')
agent5=agent5.groupby(['Region-User Type']).count()
agent5=agent5.reset_index()
agent5=agent5.rename(columns={'Store Visit Name':'TC Visits'})
Format=pd.merge(Format,agent5,on='Region-User Type',how='left')
Format['TC Achieved %']=(Format['TC Visits']/Format['TC Target'])*100
Format['TC Achieved %']=Format['TC Achieved %'].fillna(0)
Format['TC Achieved %']=Format['TC Achieved %'].round(2)
Format['TC Achieved %']=Format['TC Achieved %'].astype(str)+'%'
Format['PC Target']=Format['User Id']*25
agent6=agent3[['Region-User Type','Store Visit Order Name']]
agent6['Region-User Type']=agent6['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
agent6=agent6.drop_duplicates(subset=['Store Visit Order Name'],keep='first')
agent6=agent6[agent6['Store Visit Order Name']!=0]
agent6=agent6.groupby(['Region-User Type']).count()
agent6=agent6.reset_index()
agent6=agent6.rename(columns={'Store Visit Order Name':'PC Visits'})
Format=pd.merge(Format,agent6,on='Region-User Type',how='left')
Format['PC Achieved %']=(Format['PC Visits']/Format['PC Target'])*100
Format['PC Achieved %']=Format['PC Achieved %'].fillna(0)
Format['PC Achieved %']=Format['PC Achieved %'].round(2)
Format['PC Achieved %']=Format['PC Achieved %'].astype(str)+'%'
agent3=agent3.drop(['Region-User Type'],axis=1)
Monthly=pd.read_excel('Monthly_Agent.xlsx')
Monthly['Check In']=Monthly['Check In'].astype(str)
Monthly['Check1']=Monthly['Check In']+Monthly['Store Visitor Name']+Monthly['Store Visit Name']+Monthly['Retailer City']
agent3['Check In']=agent3['Check In'].astype(str)
agent3['Check']=agent3['Check In']+agent3['Store Visitor Name']+agent3['Store Visit Name']+agent3['Retailer City']
Monthly1=Monthly[['Check1']]
Monthly=Monthly.drop(['Check1'],axis=1)
agent3=pd.merge(agent3,Monthly1,left_on='Check',right_on='Check1',how='left')
agent3['Check1']=agent3['Check1'].fillna(0)
agent3=agent3[agent3['Check1']==0]
agent3=agent3.drop(['Check','Check1'],axis=1)
Monthly=pd.concat([Monthly,agent3],axis=0)
Monthly.to_excel('Monthly_Agent.xlsx',index=False)
Monthly['Region-User Type']=Monthly['Region']+Monthly['User Type']
Monthly1=Monthly[['Region-User Type','Total Liter']]
Monthly1['Region-User Type']=Monthly1['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
Monthly1=Monthly1.groupby(['Region-User Type']).sum()
Monthly1=Monthly1.reset_index()
Format=pd.merge(Format,Monthly1,on='Region-User Type',how='left')
Format=Format.rename(columns={'Total Liter':'Monthly Liter Sold'})
Monthly2=Monthly[['Region-User Type','Store Visit Name']]
Monthly2['Region-User Type']=Monthly2['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
Monthly2=Monthly2.drop_duplicates(subset=['Store Visit Name'],keep='first')
Monthly2=Monthly2.groupby(['Region-User Type']).count()
Monthly2=Monthly2.reset_index()
Monthly2=Monthly2.rename(columns={'Store Visit Name':'Monthly TC Visits'})
Format=pd.merge(Format,Monthly2,on='Region-User Type',how='left')
Monthly3=Monthly[['Region-User Type','Store Visit Order Name']]
Monthly3['Region-User Type']=Monthly3['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
Monthly3=Monthly3.drop_duplicates(subset=['Store Visit Order Name'],keep='first')
Monthly3=Monthly3[Monthly3['Store Visit Order Name']!=0]
Monthly3=Monthly3.groupby(['Region-User Type']).count()
Monthly3=Monthly3.reset_index()
Monthly3=Monthly3.rename(columns={'Store Visit Order Name':'Monthly PC Visits'})
a=Monthly3['Monthly PC Visits'].sum()
Format=pd.merge(Format,Monthly3,on='Region-User Type',how='left')
#Format.to_excel('test.xlsx',index=False)
Format=Format.fillna(0)
Format['Liter Sold']=Format['Liter Sold'].round(0)
Format['Monthly Liter Sold']=Format['Monthly Liter Sold'].round(0)
Final=Format
Final=Final[['Region','User Id','Working People','TC Target','Liter Sold','TC Visits','TC Achieved %','PC Target','PC Visits','PC Achieved %','Monthly Liter Sold','Monthly TC Visits','Monthly PC Visits']]
Final=Final.groupby(['Region']).sum()
Final=Final.reset_index()
Final['TC Achieved %']=(Final['TC Visits']/Final['TC Target'])*100
Final['TC Achieved %']=Final['TC Achieved %'].fillna(0)
Final['TC Achieved %']=Final['TC Achieved %'].round(2)
Final['TC Achieved %']=Final['TC Achieved %'].astype(str)+'%'
Final['PC Achieved %']=(Final['PC Visits']/Final['PC Target'])*100
Final['PC Achieved %']=Final['PC Achieved %'].fillna(0)
Final['PC Achieved %']=Final['PC Achieved %'].round(2)
Final['PC Achieved %']=Final['PC Achieved %'].astype(str)+'%'
Final['Region-User Type']=Final['Region'].replace({'Chambal':'Chambal Total','East':'East Total','North 1':'North 1 Total','North 2':'North 2 Total','South':'South Total','West':'West Total'})
Final=Final.rename(columns={'Region-User Type':'User Type'})
Final['User Type']=Final['User Type'].replace({'Chambal Total':'Total Chambal','East Total':'Total East','North 1 Total':'Total North 1','North 2 Total':'Total North 2','South Total':'Total South','West Total':'Total West'})
#Final.to_excel('Check.xlsx',index=False)
#Final=pd.read_excel('Check.xlsx')
#Final1=pd.read_excel('test.xlsx')
Final=pd.concat([Format,Final],axis=0)
Final=Final[['Region','RBM','User Type','Region-User Type','User Id','Working People','TC Target','TC Visits','TC Achieved %','PC Target','PC Visits','PC Achieved %','Liter Sold','Monthly TC Visits','Monthly PC Visits','Monthly Liter Sold']]
Final=Final.sort_values(["Region","User Type"], ascending = True)
Final.to_excel('Final.xlsx',index=False)
Final1=Final[Final['User Type']=='DSM']
Final1=Final1[['User Type','User Id','Working People','TC Target','Liter Sold','TC Visits','PC Target','PC Visits','Monthly Liter Sold','Monthly TC Visits','Monthly PC Visits']]
Final1=Final1.groupby(['User Type']).sum()
Final1=Final1.reset_index()
Final1['TC Achieved %']=(Final1['TC Visits']/Final1['TC Target'])*100
Final1['TC Achieved %']=Final1['TC Achieved %'].fillna(0)
Final1['TC Achieved %']=Final1['TC Achieved %'].round(2)
Final1['TC Achieved %']=Final1['TC Achieved %'].astype(str)+'%'
Final1['PC Achieved %']=(Final1['PC Visits']/Final1['PC Target'])*100
Final1['PC Achieved %']=Final1['PC Achieved %'].fillna(0)
Final1['PC Achieved %']=Final1['PC Achieved %'].round(2)
Final1['PC Achieved %']=Final1['PC Achieved %'].astype(str)+'%'
Final1['Region']='Zone1 Total'
Final2=Final[Final['User Type']=='Field Force']
Final2=Final2[['User Type','User Id','Working People','TC Target','Liter Sold','TC Visits','PC Target','PC Visits','Monthly Liter Sold','Monthly TC Visits','Monthly PC Visits']]
Final2=Final2.groupby(['User Type']).sum()
Final2=Final2.reset_index()
Final2['TC Achieved %']=(Final2['TC Visits']/Final2['TC Target'])*100
Final2['TC Achieved %']=Final2['TC Achieved %'].fillna(0)
Final2['TC Achieved %']=Final2['TC Achieved %'].round(2)
Final2['TC Achieved %']=Final2['TC Achieved %'].astype(str)+'%'
Final2['PC Achieved %']=(Final2['PC Visits']/Final2['PC Target'])*100
Final2['PC Achieved %']=Final2['PC Achieved %'].fillna(0)
Final2['PC Achieved %']=Final2['PC Achieved %'].round(2)
Final2['PC Achieved %']=Final2['PC Achieved %'].astype(str)+'%'
Final2['Region']='Zone2 Total'
Final3=Final
Final3['User Type']=Final3['User Type'].replace({'Total Chambal':'Total','Total East':'Total','Total North 1':'Total','Total North 2':'Total','Total South':'Total','Total West':'Total'})
Final3=Final3[Final3['User Type']=='Total']
Final3=Final3[['User Type','User Id','Working People','TC Target','Liter Sold','TC Visits','PC Target','PC Visits','Monthly Liter Sold','Monthly TC Visits','Monthly PC Visits']]
Final3=Final3.groupby(['User Type']).sum()
Final3=Final3.reset_index()
Final3['TC Achieved %']=(Final3['TC Visits']/Final3['TC Target'])*100
Final3['TC Achieved %']=Final3['TC Achieved %'].fillna(0)
Final3['TC Achieved %']=Final3['TC Achieved %'].round(2)
Final3['TC Achieved %']=Final3['TC Achieved %'].astype(str)+'%'
Final3['PC Achieved %']=(Final3['PC Visits']/Final3['PC Target'])*100
Final3['PC Achieved %']=Final3['PC Achieved %'].fillna(0)
Final3['PC Achieved %']=Final3['PC Achieved %'].round(2)
Final3['PC Achieved %']=Final3['PC Achieved %'].astype(str)+'%'
Final3['Region']='Zone3 Total'
Final0=pd.concat([Final,Final1,Final2,Final3],axis=0)
Final0=Final0.rename(columns={'User Id':'People Count'})
Final0.to_excel('Final.xlsx',index=False)