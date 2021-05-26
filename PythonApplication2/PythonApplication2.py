

import pandas as pd
import numpy as np
workDir = "C:/Users/MRehman/Desktop/Premium Reserves/"
saveDir = "C:/Users/MRehman/Desktop/Premium Reserves/Output/"


pd.set_option('display.expand_frame_repr', False)

#as
Fac = pd.read_csv(workDir + '_4_FacPremium.csv')
PT = pd.read_csv(workDir + '_4_TPPremium.csv')
NP = pd.read_csv(workDir + '_4_TNPPremium.csv')



Fac = Fac.rename(columns= {'Policy No.':'Policy No','Insured':'Insured/CedingCompany','CURRENCY DESC':'CURRENCY','TreatyPeriod':'Duration','Policy INCEPTION DATE':'Inception Date','Policy EXPIRY DATE':'Expiry Date'})
PT = PT.rename(columns= {'TREATY#':'Policy No','Ceding company':'Insured/CedingCompany','CURRENCY NAME':'CURRENCY','TreatyPeriod':'Duration'})
NP = NP.rename(columns= {'TREATY#':'Policy No','Ceding company':'Insured/CedingCompany','CURRENCY NAME':'CURRENCY','TreatyPeriod':'Duration'})
#check
Fac['Inception Date']= pd.to_datetime(Fac['Inception Date']).dt.date
PT['Inception Date']= pd.to_datetime(PT['Inception Date']).dt.date
NP['Inception Date']= pd.to_datetime(NP['Inception Date']).dt.date
Fac['Expiry Date']= pd.to_datetime(Fac['Expiry Date']).dt.date
PT['Expiry Date']= pd.to_datetime(PT['Expiry Date']).dt.date
NP['Expiry Date']= pd.to_datetime(NP['Expiry Date']).dt.date
#asd
Fac = Fac.loc[:,['Policy No','Type','Insured/CedingCompany','CURRENCY','Country','Inception Date','Expiry Date','OUTWARD TREATY NO.','UwYr','ReservingClass','RiskShare','Duration','Gr_Signed_Prem','Gr_Signed_Prem','Gr_Booked_Prem','Gr_PipePr','Pr_Signed_Prem','Gr_Written_Prem','Gr_Booked_Ded','Gr_Signed_Ded','Gr_PipeDed','Gr_Written_Ded','UnearnedDays','UPRPerc','Gr_UPR','Gr_DAC','Pr_PipePr','Pr_Booked_Prem','Pr_Written_Prem','Pr_UPR','OR COMM','Pr_PipeDed','Pr_Booked_Ded','Pr_Written_Ded','Pr_DAC','Gr_Earned_Prem','Gr_Earned_Ded','Pr_Earned_Prem']]
PT = PT.loc[:,['Policy No','Type','Insured/CedingCompany','CURRENCY','Country','Inception Date','Expiry Date','OUTWARD TREATY NO.','UwYr','ReservingClass','RiskShare','Duration','Gr_Signed_Prem','Gr_Signed_Prem','Gr_Booked_Prem','Gr_PipePr','Pr_Signed_Prem','Gr_Written_Prem','Gr_Booked_Ded','Gr_Signed_Ded','Gr_PipeDed','Gr_Written_Ded','UnearnedDays','UPRPerc','Gr_UPR','Gr_DAC','Pr_PipePr','Pr_Booked_Prem','Pr_Written_Prem','Pr_UPR','OR COMM','Pr_PipeDed','Pr_Booked_Ded','Pr_Written_Ded','Pr_DAC','Gr_Earned_Prem','Gr_Earned_Ded','Pr_Earned_Prem']]
NP = NP.loc[:,['Policy No','Type','Insured/CedingCompany','CURRENCY','Country','Inception Date','Expiry Date','OUTWARD TREATY NO.','UwYr','ReservingClass','RiskShare','Duration','Gr_Signed_Prem','Gr_Signed_Prem','Gr_Booked_Prem','Gr_PipePr','Pr_Signed_Prem','Gr_Written_Prem','Gr_Booked_Ded','Gr_Signed_Ded','Gr_PipeDed','Gr_Written_Ded','UnearnedDays','UPRPerc','Gr_UPR','Gr_DAC','Pr_PipePr','Pr_Booked_Prem','Pr_Written_Prem','Pr_UPR','OR COMM','Pr_PipeDed','Pr_Booked_Ded','Pr_Written_Ded','Pr_DAC','Gr_Earned_Prem','Gr_Earned_Ded','Pr_Earned_Prem']]

Data= pd.concat([Fac,PT,NP], ignore_index=True)
Data['Outward']=Data['OUTWARD TREATY NO.'].str[0:1]

DataPivot=pd.pivot_table(Data, values=('Gr_Written_Prem','Pr_Written_Prem'), index=['Type','ReservingClass','UwYr'],columns=['Outward'], aggfunc=np.sum)
DataPivot=pd.DataFrame(DataPivot)
DataPivot=DataPivot.reset_index()
DataPivot.columns = ['_'.join(col).strip() for col in DataPivot.columns.values]
DataPivot=DataPivot.fillna(0)
DataPivot['Gr_Written_Prem']=DataPivot['Gr_Written_Prem_N']+DataPivot['Gr_Written_Prem_O']+DataPivot['Gr_Written_Prem_R']

DataPivot=DataPivot.loc[:,['Type_','ReservingClass_','UwYr_','Gr_Written_Prem','Pr_Written_Prem_N','Pr_Written_Prem_O','Pr_Written_Prem_R']]

writer = pd.ExcelWriter(saveDir+'/Premium.xlsx', engine='xlsxwriter')
Data.to_excel(writer, sheet_name='Prem_Q1-21')
DataPivot.to_excel(writer, sheet_name='Pivot')
writer.save()
