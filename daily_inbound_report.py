import numpy as np
import pandas as pd
import datetime
import os
import openpyxl
from pandas.core.indexes.base import Index


#Sorting IR - Received
IRR = pd.read_excel('Inbound Report R.xlsx')
IRR = IRR[['ASNNO','EXPECTEDQTY','RECEIVINGTIME','RECEIVEDQTY']]
IRR ['RECEIVINGTIME'] = pd.to_datetime(IRR['RECEIVINGTIME'].str.split(',').str[0])
IRR ['Total Received'] =  IRR.groupby(['ASNNO'])['RECEIVEDQTY'].transform('sum')
IRR ['Expected QTY'] = IRR.groupby(['ASNNO'])['EXPECTEDQTY'].transform('sum')
C_IRR = IRR.sort_values(by=['RECEIVINGTIME']).drop_duplicates(subset='ASNNO', keep='first')
C_IRR = C_IRR[['ASNNO','Expected QTY','RECEIVINGTIME','Total Received']]

#Sorting IR - Putaway
IRP = pd.read_excel('Inbound Report R.xlsx')
IRP = IRP[['ASNNO','PUTAWAYTIME','PUTAWAYQTY']]
IRP ['PUTAWAYTIME'] = pd.to_datetime(IRP['PUTAWAYTIME'].str.split(',').str[0])
IRP ['Total Putaway'] =  IRP.groupby(['ASNNO'])['PUTAWAYQTY'].transform('sum')
C_IRP = IRP.sort_values(by=['PUTAWAYTIME']).drop_duplicates(subset='ASNNO', keep='first')
C_IRP = C_IRP[['ASNNO','PUTAWAYTIME','Total Putaway']]

#Finishing IR
M_IR = C_IRR.merge(C_IRP,on=['ASNNO'],how='left')
M_IR = M_IR.rename(columns={'ASNNO': 'ASN No'})

#Filling Type
IS = pd.read_csv('Inbound Schedule.csv')
IS = IS.rename(columns={
    'asnno': 'ASN No', 
    'invoicenum': 'Invoice Num',
    'customername': 'Seller Name',
    'asnstatus': 'ASN Status',
    'expectedarrivetime': 'Expected Arrival Time',
    'actualarrivetime': 'Actual Arrival Time',
    'slabreachdate': 'SLA Breach Time',    
    })
IS['Type'] = np.where(IS['Invoice Num'].str.contains("NON-BUNDLING", case=False, na=False), 
'Non Bundling', np.where(IS['Invoice Num'].str.contains("BUNDLING", case=False, na=False),'Bundling',
'Non Bundling'))

#Merge IR - IS
M_IS = IS.merge(M_IR, on='ASN No', how='left')
M_IS = M_IS[[
    'ASN No',
    'Invoice Num',
    'Seller Name',
    'ASN Status',
    'Expected Arrival Time',
    'Actual Arrival Time',
    'SLA Breach Time',
    'Expected QTY',
    'RECEIVINGTIME',
    'Total Received',
    'PUTAWAYTIME',
    'Total Putaway',
    'Type']]

conditions = [   
    (M_IS['ASN Status']=='Order Created') & (M_IS['Total Received']==0) & (M_IS['Total Putaway']==0) | 
    (M_IS['ASN Status']=='ASN Closed') & (M_IS['Total Received']==0) & (M_IS['Total Putaway']==0),
    (M_IS['Type']=='Bundling') & (M_IS['Total Received']!= M_IS['Total Putaway']),
    (M_IS['ASN Status']=='ASN Closed') & (M_IS['Total Received']== M_IS['Total Putaway']) & 
    (M_IS['RECEIVINGTIME'].notna()) & (M_IS['PUTAWAYTIME'].notna()) & (M_IS['PUTAWAYTIME'] < M_IS['SLA Breach Time'])|
    (M_IS['ASN Status']=='Fully Received') & (M_IS['Total Received']== M_IS['Total Putaway']) & 
    (M_IS['Total Received']!= 0) & (M_IS['Total Putaway']!= 0) &
    (M_IS['RECEIVINGTIME'].notna()) & (M_IS['PUTAWAYTIME'].notna()) & (M_IS['PUTAWAYTIME'] < M_IS['SLA Breach Time'])|
    (M_IS['ASN Status']=='Partially Received') & (M_IS['Total Received']== M_IS['Total Putaway']) &
    (M_IS['Total Received']!= 0) & (M_IS['Total Putaway']!= 0) &
    (M_IS['RECEIVINGTIME'].notna()) & (M_IS['PUTAWAYTIME'].notna()) & (M_IS['PUTAWAYTIME'] < M_IS['SLA Breach Time'])
    ]
caveat = ['Not Coming','On Process','Achieved']
M_IS['SLA'] = np.select(conditions, caveat, default='Check')

M_IS.to_excel('Inbound Report Trial.xlsx', index=False)

#pd.set_option('display.max_columns', None)
#print(M_IS)




