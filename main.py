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


reco_file_one_path=os.path.join(cwd,"temp","gst_reco_completed.csv")
reco_file_two_path=os.path.join(cwd,"temp","gst_reco.csv")

if os.path.exists(reco_file_one_path) and os.path.exists(reco_file_two_path):
    os.remove(reco_file_one_path)
    os.remove(reco_file_two_path)

for key,value in company_gst_master.items():
    
    company_df=pd.DataFrame([key,value]).transpose()
    account_temp_df=accounts_df.loc[accounts_df['Sales Tax No.'] == key]
    acc_df=pd.DataFrame(columns=["Source","Date","Narration","Gross Total"])
    
    for index,row in account_temp_df.iterrows():       
        temp_df = {"Source": "Books", 'Date': row["Date"],"Narration":row["Narration"][8:].strip(),"Gross Total":row["Gross Total"]}
        acc_df=acc_df.append(temp_df,ignore_index=True)

    gstr_temp_df=gstr_df.loc[gstr_df['GSTIN of supplier']==key]
    gst_df=pd.DataFrame(columns=["GST Source","Invoice Date","Invoice number","Invoice Value"])
    for index,row in gstr_temp_df.iterrows():
        row["Invoice number"]=row["Invoice number"].strip()
        temp_df={"GST Source":"GSTR2B Input","Invoice Date":row["Invoice Date"],"Invoice number":row["Invoice number"],"Invoice Value":row["Invoice Value"]}
        gst_df=gst_df.append(temp_df,ignore_index=True)
    
    acc_gst_merge=pd.merge(acc_df,gst_df,left_on="Narration",right_on="Invoice number",how="outer")

    if acc_gst_merge["Source"].isnull().values.any() or acc_gst_merge["GST Source"].isnull().values.any(): 
        company_df.to_csv(reco_file_two_path, mode='a',header=False,index=False)
        acc_gst_merge.to_csv(reco_file_two_path, mode='a',index=False) 
    else:
        company_df.to_csv(reco_file_one_path, mode='a',header=False,index=False)
        acc_gst_merge.to_csv(reco_file_one_path, mode='a',index=False)


os.startfile(reco_file_one_path)
os.startfile(reco_file_two_path)
