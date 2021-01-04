import openpyxl as op
import os
import pandas as pd
import xlrd
import shutil

cwd=os.getcwd()
accounts_file_path=os.path.join(cwd,"data","FA-SALES-Nov-20.xlsx")
gstr_file_path=os.path.join(cwd,"data","FA-GSTR2b-Nov-20.xlsx")

accounts_df=pd.DataFrame()
gstr_df=pd.DataFrame()

if os.path.exists(accounts_file_path):
    accounts_wb=op.load_workbook(accounts_file_path)
    accounts_df=pd.DataFrame(accounts_wb.active.values)
    accounts_df.columns=accounts_df.iloc[2]
    accounts_df=accounts_df[3:-1]
else:
    print("Accounts file missing")


if os.path.exists(gstr_file_path):
    gstr_wb=op.load_workbook(gstr_file_path)
    gstr_df=pd.DataFrame(gstr_wb.active.values)
    gstr_df.columns=["GSTIN of supplier","Trade/Legal name","Invoice number","Invoice type","Invoice Date",	"Invoice Value","Place of supply",
    	"Supply Attract Reverse Charge","Rate","Taxable Value","Integrated Tax","Central Tax","State/UT Tax","Cess",
        "GSTR-1/5 Period","GSTR-1/5 Filing Date","ITC Availability","Reason","Applicable Tax Rate","Source","IRN","IRN Date"]
    gstr_df=gstr_df[6:]
else:
    print("GSTR file missing")

company_gst_master=dict()

for index, row in accounts_df.iterrows():
    company_gst_master[row['Sales Tax No.']]=row['Particulars']


temp_file_path=os.path.join(cwd,"temp","gst_reco.csv")

if os.path.exists(temp_file_path):
    os.remove(temp_file_path)

for key,value in company_gst_master.items():
    company_df=pd.DataFrame([key,value]).transpose()
    company_df.to_csv(temp_file_path, mode='a',header=False,index=False)
    temp_df=pd.DataFrame(["Input","Date","Bill Number","Total"]).transpose()
    temp_df.to_csv(temp_file_path, mode='a',header=False,index=False)

    account_temp_df=accounts_df.loc[accounts_df['Sales Tax No.'] == key]
    for index,row in account_temp_df.iterrows():
        temp_df=pd.DataFrame(["Tally", row["Date"], row["Narration"],row["Gross Total"]]).transpose()
        temp_df.to_csv(temp_file_path, mode='a',header=False,index=False)
    
    gstr_temp_df=gstr_df.loc[gstr_df['GSTIN of supplier']==key]
    for index,row in gstr_temp_df.iterrows():
        temp_df=pd.DataFrame(["GST Input",row["Invoice Date"], row["Invoice number"],row["Invoice Value"]]).transpose()
        temp_df.to_csv(temp_file_path, mode='a',header=False,index=False)

os.startfile(temp_file_path)