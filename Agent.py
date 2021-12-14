import pandas as pd
import numpy as np
master=pd.read_excel('Master.xlsx')
U_Master=pd.read_excel('User Master.xlsx')
master=master.drop_duplicates(subset=['Material'],keep='first')
agent=pd.read_excel('B2C_agent.xlsx')
agent=pd.merge(agent,master,left_on='Product: SAP Number',right_on='Material',how='left')
agent=agent[['Check In','Owner Business Line','Created By: Full Name','Store Visit Name','Store Visit Order Name','Store Visit Order Product Name','Status','No Order Reason','Product: Brand','Product: SAP Number','Product: Product Name','Remote Visit','Account: Account Name','Account: City: Store Visit Related Info Name','Owner Business Zone','Account: State','Account: Region','Created By: Manager: Full Name','Created By: Manager: Manager: Full Name','Designation','Distributor: SAP Account Number','Distributor: Account Name','Beat','Fulfillment','Order Quantity (Case)','Order Quantity (Piece)','Gross Weight','Pack size']]
agent=agent.fillna(0)
agent1=agent[agent['Order Quantity (Piece)']>0]
agent2=agent[agent['Order Quantity (Piece)']==0]
agent1['case']=agent1['Order Quantity (Piece)']/agent1['Pack size']
agent1=agent1.drop(['Order Quantity (Case)','Order Quantity (Piece)'],axis=1)
agent1=agent1.rename(columns={'case':'Order Quantity (Case)'})
agent2=agent2.drop(['Order Quantity (Piece)'],axis=1)
agent3=pd.concat([agent1,agent2],axis=0)
agent3['Total Liter']=agent3['Gross Weight']*agent3['Order Quantity (Case)']
SVO=pd.read_excel('SVO.xlsx')
agent3=pd.merge(agent3,SVO,left_on='Store Visit Name',right_on='Name',how='left')
agent3=pd.merge(agent3,U_Master,left_on='LastModifiedById',right_on='User ID',how='left')
agent6=agent3.drop(['Name','CreatedById','LastModifiedById','Full Name','Roll'],axis=1)
agent6['RBM']=agent6['Region'].replace({'East':'Arun Neogi','Chambal':'Mahesh Kumar','North 1':'Sandeep Kaul','North 2':'Rohit Nair','South':'RS Murthy','West':'Naresh Makhija'})
agent6=agent6.sort_values(['Created By: Full Name','Store Visit Order Name'])
agent6.to_excel('daily.xlsx',index=False)
Format=pd.read_excel('Format.xlsx')
Format['Region-User Type']=Format['Region']+Format['User Type']
agent6['Region-User Type']=agent6['Region']+agent6['User Type']
U_Master['Region-User Type']=U_Master['Region']+U_Master['User Type']
U_Master1=U_Master[U_Master['Biz']=='B2C']
U_Master1=U_Master1[U_Master1['User Type'].isin(['Field Force','DSM'])]
U_Master2=U_Master1[['Full Name','Region-User Type']]
U_Master2=U_Master2.groupby(['Region-User Type']).count()
U_Master2=U_Master2.reset_index()
Format=pd.merge(Format,U_Master2,on='Region-User Type',how='left')
Format['TC Target']=Format['Full Name']*40

agent7=agent6[['Region-User Type','Total Liter']]
agent7['Region-User Type']=agent7['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
agent7=agent7.groupby(['Region-User Type']).sum()
agent7=agent7.reset_index()
Format=pd.merge(Format,agent7,on='Region-User Type',how='left')
Format=Format.rename(columns={'Total Liter':'Liter Sold'})

agent8=agent6[['Region-User Type','Store Visit Name']]
agent8['Region-User Type']=agent8['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
agent8=agent8.drop_duplicates(subset=['Store Visit Name'],keep='first')
agent8=agent8.groupby(['Region-User Type']).count()
agent8=agent8.reset_index()
agent8=agent8.rename(columns={'Store Visit Name':'TC Visits'})
Format=pd.merge(Format,agent8,on='Region-User Type',how='left')

Format['TC Achieved %']=(Format['TC Visits']/Format['TC Target'])*100
Format['TC Achieved %']=Format['TC Achieved %'].fillna(0)
Format['TC Achieved %']=Format['TC Achieved %'].round(2)
Format['TC Achieved %']=Format['TC Achieved %'].astype(str)+'%'
Format['PC Target']=Format['Full Name']*25

agent9=agent6[['Region-User Type','Store Visit Order Name']]
agent9['Region-User Type']=agent9['Region-User Type'].replace({'EastASM':'EastField Force','North 1ASM':'North 1Field Force','North 2ASM':'North 2Field Force','ChambalASM':'ChambalField Force','WestASM':'WestField Force','SouthASM':'SouthField Force'})
agent9=agent9.drop_duplicates(subset=['Store Visit Order Name'],keep='first')
agent9=agent9[agent9['Store Visit Order Name']!=0]
agent9=agent9.groupby(['Region-User Type']).count()
agent9=agent9.reset_index()
agent9=agent9.rename(columns={'Store Visit Order Name':'PC Visits'})
Format=pd.merge(Format,agent9,on='Region-User Type',how='left')

Format['PC Achieved %']=(Format['PC Visits']/Format['PC Target'])*100
Format['PC Achieved %']=Format['PC Achieved %'].fillna(0)
Format['PC Achieved %']=Format['PC Achieved %'].round(2)
Format['PC Achieved %']=Format['PC Achieved %'].astype(str)+'%'

agent6=agent6.drop(['Region-User Type'],axis=1)
Monthly=pd.read_excel('Monthly_Agent.xlsx')
Monthly['Check In']=Monthly['Check In'].astype(str)
Monthly['Check1']=Monthly['Check In']+Monthly['Created By: Full Name']+Monthly['Store Visit Name']+Monthly['Account: City: Store Visit Related Info Name']
agent6['Check In']=agent6['Check In'].astype(str)
agent6['Check']=agent6['Check In']+agent6['Created By: Full Name']+agent6['Store Visit Name']+agent6['Account: City: Store Visit Related Info Name']
Monthly1=Monthly[['Check1']]
Monthly=Monthly.drop(['Check1'],axis=1)
agent6=pd.merge(agent6,Monthly1,left_on='Check',right_on='Check1',how='left')
agent6['Check1']=agent6['Check1'].fillna(0)
agent6=agent6[agent6['Check1']==0]
agent6=agent6.drop(['Check','Check1'],axis=1)
Monthly=pd.concat([Monthly,agent6],axis=0)
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
Format=Format.fillna(0)
Format['Liter Sold']=Format['Liter Sold'].round(0)
Format['Monthly Liter Sold']=Format['Monthly Liter Sold'].round(0)
Format.to_excel('test.xlsx',index=False)
Final=Format
Final=Final[['Region','Full Name','TC Target','Liter Sold','TC Visits','TC Achieved %','PC Target','PC Visits','PC Achieved %','Monthly Liter Sold','Monthly TC Visits','Monthly PC Visits']]
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
Final.to_excel('Check.xlsx',index=False)
Final=pd.read_excel('Check.xlsx')
Final1=pd.read_excel('test.xlsx')
Final=pd.concat([Final,Final1],axis=0)
Final=Final[['Region','RBM','User Type','Region-User Type','Full Name','TC Target','TC Visits','TC Achieved %','PC Target','PC Visits','PC Achieved %','Liter Sold','Monthly TC Visits','Monthly PC Visits','Monthly Liter Sold']]
Final=Final.sort_values(["Region","User Type"], ascending = True)
Final.to_excel('Final.xlsx',index=False)
Final1=Final[Final['User Type']=='DSM']
Final1=Final1[['User Type','Full Name','TC Target','Liter Sold','TC Visits','PC Target','PC Visits','Monthly Liter Sold','Monthly TC Visits','Monthly PC Visits']]
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
Final2=Final2[['User Type','Full Name','TC Target','Liter Sold','TC Visits','PC Target','PC Visits','Monthly Liter Sold','Monthly TC Visits','Monthly PC Visits']]
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
Final3=Final3[['User Type','Full Name','TC Target','Liter Sold','TC Visits','PC Target','PC Visits','Monthly Liter Sold','Monthly TC Visits','Monthly PC Visits']]
Final3=Final3.groupby(['User Type']).sum()
Final3=Final3.reset_index()
Final3['TC Achieved %']=(Final3['TC Visits']/Final3['TC Target'])*100
Final3['TC Achieved %']=Final3['TC Achieved %'].fillna(0)
Final3['TC Achieved %']=Final3['TC Achieved %'].round(2)
Final3['TC Achieved %']=Final3['TC Achieved %'].astype(str)+'%'

Final3['PC Achieved %']=(Final3['PC Visits']/Final2['PC Target'])*100
Final3['PC Achieved %']=Final3['PC Achieved %'].fillna(0)
Final3['PC Achieved %']=Final3['PC Achieved %'].round(2)
Final3['PC Achieved %']=Final3['PC Achieved %'].astype(str)+'%'
Final3['Region']='Zone3 Total'
Final0=pd.concat([Final,Final1,Final2,Final3],axis=0)
Final0=Final0.rename(columns={'Full Name':'People Count'})
Final0.to_excel('Final.xlsx',index=False)
