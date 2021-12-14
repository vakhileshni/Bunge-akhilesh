import subprocess
import io
import pandas as pd
from io import StringIO
from datetime import date
from datetime import timedelta
today = date.today()
yesterday = today - timedelta(days = 1)
yesterday=str(yesterday)
a='Select Check_In__c,Check_Out__c,Id,Owner_Business_Line__c,Store_Visit_Owner__c,Name,Status__c,No_Order_Reason__c,Remote_Visit__c,Account__c,LastModifiedById FROM Store_Visit__c WHERE LastModifiedDate > '+yesterday+'T00:00:01.000Z AND LastModifiedDate < '+yesterday+'T23:59:59.000Z'
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"' -u -Pro -r csv"
p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
StringData = StringIO(p.stdout)
df = pd.read_csv(StringData, sep =",")
df=df.fillna(0)
df['Check']=df['Check_In__c'].astype('str')+df['Check_Out__c'].astype('str')
df=df[df['Check']!='00']
a='Select Id,Name,Store_Visit__c FROM Store_Visit_Order__c WHERE LastModifiedDate > '+yesterday+'T00:00:01.000Z AND LastModifiedDate < '+yesterday+'T23:59:59.000Z'
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"' -u -Pro -r csv"
p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
StringData = StringIO(p.stdout)
df1 = pd.read_csv(StringData, sep =",")
a='Select Name,Store_Visit_Order__c,Order_Quantity_Case__c,Order_Quantity_Piece__c,Product2__c,Fulfillment__c,Distributor__c FROM Store_Visit_Order_Product__c WHERE LastModifiedDate > '+yesterday+'T00:00:01.000Z AND LastModifiedDate < '+yesterday+'T23:59:59.000Z'
a=str(a)
b="sfdx force:data:soql:query -q '"+a+"' -u -Pro -r csv"
p=subprocess.run(['powershell',b],shell=True,stdout=subprocess.PIPE,text=True)
StringData = StringIO(p.stdout)
df2 = pd.read_csv(StringData, sep =",")
with io.open('Product.csv','r',encoding='utf16') as f:
    text = f.read()
StringData = StringIO(text)
pro = pd.read_csv(StringData, sep =",")
pro=pro.rename(columns={'Name':'DESCRIPTION'})
df=pd.merge(df,df1,left_on='Id',right_on='Store_Visit__c',how='left')
df=pd.merge(df,df2,left_on='Id_y',right_on='Store_Visit_Order__c',how='left')
df=df.drop(['Id_y','Id_x','Store_Visit__c','Store_Visit_Order__c'],axis=1)
df=pd.merge(df,pro,left_on='Product2__c',right_on='Id',how='left')
df=df.rename(columns={'Name_x':'Store_Visit_Name','Name_y':'Store_Visit_Order_Name','Name':'Store_Visit_Order_Product_Name','Check_In__c':'Check_In','Check_Out__c':'Check_Out','Owner_Business_Line__c':'Owner_Business_Line','Store_Visit_Owner__c':'Store_Visit_Owner','Status__c':'Status','No_Order_Reason__c':'No_Order_Reason','Remote_Visit__c':'Remote_Visit','Order_Quantity_Case__c':'Order_Quantity_Case','Order_Quantity_Piece__c':'Order_Quantity_Piece','Fulfillment__c':'Fulfillment'})
df=df[['Check_In','Check_Out','Owner_Business_Line','Store_Visit_Owner','Store_Visit_Name','Store_Visit_Order_Name','Store_Visit_Order_Product_Name','Status','No_Order_Reason','Remote_Visit','Account__c','LastModifiedById','Order_Quantity_Case','Order_Quantity_Piece','ProductCode','DESCRIPTION','Fulfillment','Distributor__c']]
ak=pd.read_csv(r"C:\Users\UAKHPAL\Downloads\Retailer.csv")
dfr=df[['Account__c']]
dfr=dfr.drop_duplicates(subset='Account__c',keep='first')
ak=ak[['ID','BIZ_DEFINED_GEOZONE__C','CITY_NAME__C','MOBILE_PHONE__C','NAME','OWNER_MANAGER__C','RETAILER_REGION__C']]
ak=ak.drop_duplicates(subset='ID',keep='first')
dfr=pd.merge(dfr,ak,left_on='Account__c',right_on='ID',how='left')
df=pd.merge(df,dfr,on='Account__c',how='left')
dfd=df[['Distributor__c']]
dfd=dfd.drop_duplicates(subset='Distributor__c',keep='first')
ak1=ak[['ID','NAME']]
dfd=pd.merge(dfd,ak1,left_on='Distributor__c',right_on='ID',how='left')
df=pd.merge(df,dfd,on='Distributor__c',how='left')
df.to_excel('total final.xlsx',index=False)
