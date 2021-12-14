import pandas as pd
import os
a='C:/Users/UAKHPAL/OneDrive - BUNGE/Desktop/akhilesh/Beat'
beat=pd.read_excel('beat.xlsx')
files = os.listdir('C:/Users/UAKHPAL/OneDrive - BUNGE/Desktop/akhilesh/Beat')
for name in files:
    ak=pd.read_excel(a+"/"+name)
    beat=pd.concat([beat,ak])
beat.to_excel('Beat_Plan.xlsx',index=False)