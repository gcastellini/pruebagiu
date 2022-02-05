# -*- coding: utf-8 -*-
"""
Spyder Editor
This is a temporary script file. Hago una pruebita
"""
import pandas as pd
import datetime as dt
from datetime import timedelta
date=dt.date.today()
year=str(date.year)
month=str(date.month)
day=str(date.day)
"""
User Input
"""
invfile=input("Path to Invoices.csv:  ")
payfile=input("Path to Payments.csv:  ")
difffile=input("Path to Diff Payment Date vs Due Date.csv: ")
collecfile=input("Path to Collections.csv: ")
lastres=input("Path to Last Month Results file: ")
summaryfile=input('Path to results: ')

#Files
invoices=pd.read_csv(invfile)
payments=pd.read_csv(payfile)
diffdate=pd.read_csv(difffile)
collections=pd.read_csv(collecfile)
lastresults=pd.DataFrame()
def read():
    global lastresults
    if lastres != '':
      lastresults=pd.read_excel(lastres,sheet_name='Results')
read()
lastchart=pd.DataFrame()
def readchart():
    global lastchart
    if lastres != '':
      lastchart=pd.read_excel(lastres,sheet_name='Cobrado Mes a Mes')
readchart()

#Clean Invoices
invoices=invoices[invoices.Step != "Invoice Fully Paid"]
invoices=invoices[invoices.Step != "Invoice Void"]
invoices=invoices[invoices.Step != "Invoice Closed"]
invoices=invoices[invoices.Step != "Invoice Write Off"]
invoices=invoices[invoices.BalanceUSD != 0]

#Clean Collections
collections=collections[['Ref.#','Date.1','Type DS','Feedback DS','Estimated Collection Date','Comments']]
collections2=collections.drop_duplicates('Ref.#',keep='last')

#Merge with Payments,Collections and Diff Date
payments1=payments[['Pipeline','Step']]
collections2=collections2.rename(index=str,columns={'Ref.#':'Pipeline'})
collections3=collections2[['Pipeline','Type DS','Feedback DS','Estimated Collection Date']]
diffdate1=diffdate[['Name','PromDiffTable']]
merge1=pd.merge(invoices,collections3,how='left',on='Pipeline')
merge1=merge1.fillna(0)
merge2=pd.merge(merge1,payments1,how='left',on='Pipeline')
merge2=merge2.rename(index=str,columns={'Step_x':'Step', 'Step_y':'Payment'})
merge2=merge2.fillna(0)
merge3=pd.merge(merge2,diffdate1,how='left',on='Name')
merge3=merge3.fillna(0)

#Dates
filtered=merge3.sort_values(by ='Invoice Due Date')
filtered['Invoice Due Date']=pd.to_datetime(filtered['Invoice Due Date'])

#Filter
filtered['Status']=filtered['Invoice Due Date'].apply(lambda x: 'Overdue' if x < date else 'Not Overdue')
overdue=filtered.set_index('Status')
overdue=overdue.loc[['Overdue']]
overdue=overdue.reset_index()
overdue['Status']=overdue['Feedback DS'].apply(lambda x:'Promise' if x =='Answer Received  - Payment Promise' else 'Overdue - No Promise')

##Promise
promise=overdue.set_index('Status')
promise=promise.loc[['Promise']]
promise=promise.reset_index()
promise['Status']=promise['Estimated Collection Date']
promise['Status']=promise['Estimated Collection Date'].apply(lambda x:'Overdue' if x==0 else x)
promise['Estimated Collection Date']=pd.to_datetime(promise['Estimated Collection Date'])
promise['Status']=promise['Estimated Collection Date'].apply(lambda x:'Overdue - Promise' if x<date else x)
promise1=promise[promise.Status != 'Overdue - Promise']
promise2=promise[promise.Status == 'Overdue - Promise']
promise2['Status']=promise2['Invoice Due Date']+pd.TimedeltaIndex(promise2['PromDiffTable'],unit='days')
promise2['Status']=promise2['Status'].apply(lambda x: 'Overdue - Promise' if x<date else x)
prom=promise2[promise2.Status != 'Overdue - Promise']
promise1=promise1.append(prom)
promise2=promise2[promise2.Status == 'Overdue - Promise']

##No Promise
overdue=overdue.set_index('Status')
overdue=overdue.loc[['Overdue - No Promise']]
overdue=overdue.reset_index()
overdue['Status']=overdue['Invoice Due Date']+pd.TimedeltaIndex(overdue['PromDiffTable'],unit='days')
overdue['Status']=overdue['Status'].apply(lambda x: 'Overdue - No Promise' if x<date else x)
overdue1=overdue[overdue.Status != 'Overdue - No Promise']
overdue=overdue[overdue.Status == 'Overdue - No Promise']

##Not Overdue
notoverdue=filtered.set_index('Status')
notoverdue=notoverdue.loc[['Not Overdue']]
notoverdue=notoverdue.reset_index()
notoverdue['Status']=notoverdue['Estimated Collection Date'].apply(lambda x:x if x !=0  else 'Not Overdue')
notoverdue['Estimated Collection Date']=pd.to_datetime(notoverdue['Estimated Collection Date'])
notoverdue1=notoverdue.set_index('Status')
notoverdue1=notoverdue1.loc[['Not Overdue']]
notoverdue1=notoverdue1.reset_index()
notoverdue1['Status']=notoverdue1['Invoice Due Date']+pd.TimedeltaIndex(notoverdue1['PromDiffTable'],unit="days")
notoverdue=notoverdue[notoverdue.Status != 'Not Overdue']
final3=notoverdue.append(notoverdue1)

##Append
final1=promise2.append(overdue)
final1['Month']=final1['Status']
final1['Period']=final1['Month']
final4=final3.append(promise1)
final2=final4.append(overdue1)
final2['Status']=pd.to_datetime(final2['Status'])
final2['Month']=final2['Status'].dt.month.astype(str)
final2['Year']=final2['Status'].dt.year.astype(str)
final2['Month']=final2['Month'].apply(lambda x: month if x < month else x)
final2['Period']=final2[["Year","Month"]].agg('-'.join,axis=1)
final=final1.append(final2)
final=final.sort_values(by=['Period'],ascending=True)

#Remove Duplicates
final=final.drop_duplicates(subset=None,keep='first',inplace=False)

#Summary
summary=final[['Period','BalanceUSD']]
summary=summary.groupby('Period').sum()
summary=summary.sort_values(by=['Period'],ascending=True)
summary.loc['Total']=summary.sum(axis=0,numeric_only=True)

#By Company
bycompany=final[['Period','Avature Company','BalanceUSD']]
bycompany=bycompany.groupby(['Period','Avature Company']).sum()
bycompany=bycompany.reset_index()
bycompany=bycompany.pivot(index='Avature Company',columns='Period',values='BalanceUSD')
bycompany=bycompany.fillna(0)
bycompany.loc['Total']=bycompany.sum(axis=0,numeric_only=True)
bycompany['Total']=bycompany.sum(axis=1)

#Last Month
finalsum=pd.DataFrame()
def last():
    global finalsum
    global lastresults
    global lastchart
    if lastres != '':
      pay=payments.rename(index=str,columns={'Payment Date':'Pay_Date'})
      pay['Pay_Date']=pd.to_datetime(pay['Pay_Date'])
      pay['last_month']=pay['Pay_Date'].apply(lambda x: 'last_month' if x > date-timedelta(days=30)   else 0)
      pay=pay[pay.last_month == 'last_month']
      lastresults1=lastresults[['Pipeline','Period']]
      pay=pay[['Name','Pipeline','Pay_Date','USDTotalInvoice','Invoice Due Date']]
      pay=pd.merge(pay,lastresults1,on='Pipeline',how='left')
      paysummary=pay[['Period','USDTotalInvoice']]
      paysummary=paysummary.rename(index=str,columns={'USDTotalInvoice':'Cobrado Mes '+month})
      paysummary=paysummary.groupby(['Period']).sum()
      paysummary=paysummary.reset_index()
      lastresults=lastresults[['Period','USDTotalInvoice']]
      lastresults=lastresults.rename(index=str,columns={'USDTotalInvoice':'Estimado Mes '+month})
      lastresults=lastresults.groupby(['Period']).sum()
      lastresults=lastresults.reset_index()
      finalsum=pd.merge(paysummary,lastresults,on=['Period'],how='outer')
      finalsum['%']=finalsum['Cobrado Mes '+month]/finalsum['Estimado Mes '+month]
      finalsum=finalsum.sort_values(by='Period')
      finalsum=finalsum.set_index('Period')
      finalsum.loc['Total']=finalsum.sum(axis=0,numeric_only=True)
      finalsum=finalsum.append(lastchart)
last()


#Export
writer=pd.ExcelWriter(summaryfile+'\Collections '+year+'-'+month+'-'+day+'.xlsx')
summary.to_excel(writer,index=True,sheet_name='Estimation '+year+'-'+month+'-'+day)
finalsum.to_excel(writer,index=True,sheet_name='Cobrado Mes a Mes')
bycompany.to_excel(writer,index=True,sheet_name='By Company')
final.to_excel(writer,index=False,sheet_name='Results')
writer.save()



#Print
print("Export Completed Successfully")



