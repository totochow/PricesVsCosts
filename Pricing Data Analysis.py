#!/usr/bin/env python
# coding: utf-8

# In[11]:


#getting all the tools ready
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import datetime

#Setting up dataframe
pricesheet = pd.DataFrame()
sales = pd.DataFrame()
mark = pd.DataFrame()
df = pd.DataFrame()

#Reading in the file:
path ='C:\\Users\\tchow\\Desktop\\Price Sheets Analysis\\Price Sheets Analysis 20200902b.xlsx'
pricesheet = pd.read_excel(path, 'Price Sheets Pricing')
sales = pd.read_excel(path, 'Sales -12 mo')
mark = pd.read_excel(path, 'Market Sheet')


# In[12]:


#setting up the item number list...
pricesheet['UOFM'] = pricesheet['UOFM'].str.upper()
df = pricesheet[['ITEMNMBR','ITEMDESC','BUYERID','Base UM','UOFM']].sort_values(by=['ITEMNMBR','Base UM','UOFM']).drop_duplicates().reset_index(drop=True)
df


# In[13]:


#sales
#pricesheet.PRCSHID.unique()

#building in all the pricesheets name into columns and adding in the details
for col in pricesheet.PRCSHID.unique():
    df = pd.merge(df,pricesheet[(pricesheet['PRCSHID'] == col)&(pricesheet['ACTIVE'] == 1)][['ITEMNMBR','Base UM','UOFM','Price']],left_on=['ITEMNMBR','Base UM','UOFM'],right_on=['ITEMNMBR','Base UM','UOFM'], how='left').rename(columns={'Price': col})


# In[14]:


#df2 = sales[['ITEMNMBR']].sort_values(by=['ITEMNMBR']).drop_duplicates().reset_index(drop=True)
#df2


#sales.pivot(index='ITEMNMBR',columns =['QTYSLCTD','EXTDCOST','XTNDPRCE'])

#adding in the ext cost, ext price, qty sold, num of unique orders the item was in, and calculating the average unit cost
sales_pivot1 = pd.pivot_table(sales,index=['ITEMNMBR'], values =['QTYSLCTD','EXTDCOST','XTNDPRCE'], aggfunc=np.sum)
sales_pivot2 = pd.pivot_table(sales,index=['ITEMNMBR'], values = ['SOPNUMBE'], aggfunc=pd.Series.nunique)

sales_merged = pd.merge(sales_pivot1,sales_pivot2,left_index=True, right_index=True)
sales_merged['avg_hist_unit_cost'] = sales_merged['EXTDCOST']/sales_merged['QTYSLCTD']


# In[15]:


MS_QOH_pivot1 = pd.pivot_table(mark,index=['ITEMNMBR'], values =['QTYAVAILB','VALUE'], aggfunc=np.sum)

MS_QOH_pivot1['avg_qoh_unit_cost'] = MS_QOH_pivot1['VALUE'] /MS_QOH_pivot1['QTYAVAILB'] 
MS_QOH_pivot1


# In[16]:


#df
#sales_merged

#getting all the details together...
temp1 = pd.merge(df,sales_merged, left_on = 'ITEMNMBR', right_index = True, how = 'left')
final = pd.merge(temp1,MS_QOH_pivot1, left_on = 'ITEMNMBR', right_index = True, how = 'left')

final


# In[17]:


final


# In[18]:


qoh_pct = 1.00

conditions = [
    pd.isna(final['avg_qoh_unit_cost']) & pd.isna(final['avg_hist_unit_cost']),
    pd.isna(final['avg_qoh_unit_cost']) & pd.notna(final['avg_hist_unit_cost']),
    pd.notna(final['avg_qoh_unit_cost']) & pd.isna(final['avg_hist_unit_cost']),
    pd.notna(final['avg_qoh_unit_cost']) & pd.notna(final['avg_hist_unit_cost'])]

choices = [
    np.nan, 
    final['avg_hist_unit_cost'], 
    final['avg_qoh_unit_cost'], 
    ((1-qoh_pct)*final['avg_hist_unit_cost']+qoh_pct*final['avg_qoh_unit_cost']) ]

final['weighted_unit_cost'] = np.select(conditions,choices, default = np.nan)
final['GM'] = (final['XTNDPRCE'] - final['EXTDCOST']) / final['XTNDPRCE']



final


# In[19]:


# create excel writer
now = datetime.datetime.now()
writer = pd.ExcelWriter('C:\\Users\\tchow\\Desktop\\results_'+now.strftime('%Y%m%d')+'.xlsx')
# write dataframe to excel sheet named 'marks'
final.to_excel(writer, 'Test')
# save the excel file
writer.save()
print('Voilà!! Successfully written to Excel Sheet.')


# In[10]:


#sales0 = sales
#sales0.set_index('CUSTNMBR').plot.pie(y='XTNDPRCE',figsize=(20,20), fontsize = 20, legend = False)
#plt.ylabel('')


# In[38]:


list(final.columns)[6:239]


# In[65]:


for k in list(final.columns)[6:236]:
    #temp = final[[k,'XTNDPRCE']]
    #print(temp[temp[k] != 'NaN'])
    print(final[final[k].notnull()][[k,'XTNDPRCE']])


# In[ ]:




