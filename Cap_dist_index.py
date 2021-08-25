# -*- coding: utf-8 -*-
"""
Created on Mon Aug 23 16:06:28 2021

@author: VR787FC
"""
import PyPDF4,os,re,pandas as pd
import pdfminer.layout
import pdfminer.high_level
from pdfminer.layout import LAParams
from io import StringIO

basedir = r'{}'.format(input("Enter the folder path: "))
filelist = [x for x in os.listdir(basedir) if ".pdf" in x]

#extract text and export to excel
all_df = pd.DataFrame(columns=['FileName','Date','LP_Name','Amount'])
for pdffile in filelist:
    output_string = StringIO()
    with open('{}/{}'.format(basedir,pdffile), 'rb') as input_pdf:
        pdfminer.high_level.extract_text_to_fp(input_pdf, output_string, laparams=LAParams())
        raw_text = output_string.getvalue().strip()
    try:
        LetterDate = re.findall('(January|February|March|April|May|June|July|August|September|October|November|December)[\s\n ]*(\d{1,2})',raw_text)[0]
    except:
        LetterDate=0
    try:        
        LP_Name = re.findall('2021[\n ]*([A-z,.\d -]*)',raw_text)[0].strip()
    except:
        LP_Name = 0
    try:
        Amount = re.findall('US\$([\s0-9,]*)[ ]*\nTotal Distribution',raw_text)[0].replace('\n','').replace(',','') #Distribution non table type 2
        # Amount = re.findall('[ \n]*\$([\s0-9,]*)[ \n]*\$[\s0-9,]*[\n0x ]*Company No',raw_text)[0].replace('\n','').replace(',','') #Distribution Total Table no (A)
        # Amount = re.findall(' \n\$([\s0-9,]*)[ ]*\(A\)',raw_text)[0].replace('\n','').replace(',','') #Distribution Total Table with (A)
        # Amount = re.findall('\nReference:[ \n]*[\s]*\$([\s0-9,]*)',raw_text)[0].replace('\n','').replace(',','') #Distribution non table type 1
        # Amount = re.findall('Your[ ]*called[ ]*capital[ ]*contribution[ ]*is[ ]*US\$([\s0-9,]*)',raw_text)[0].replace('\n','').replace(',','') #Cap_Call
    except:
        Amount=0
    df_temp = pd.DataFrame([[pdffile,LetterDate,LP_Name,Amount]],columns=['FileName','Date','LP_Name','Amount'])
    all_df = all_df.append(df_temp)
print(all_df)
#generate summary excel
all_df.to_excel('{}/summary.xlsx'.format(basedir))

input("===Press Enter twice to continue===")
input("===Press Enter once to continue===")
#rename pdf
rename_df = pd.read_excel(open('{}/summary.xlsx'.format(basedir,), 'rb'),sheet_name='Rename',engine='openpyxl')
def change_names(row):
    os.rename("{}/{}".format(basedir,row[0]), "{}/{}".format(basedir,row[1]))
print(rename_df)
input("===Press Enter twice to continue===")
input("===Press Enter once to continue===")
rename_df[['FileName', 'NewName']].apply(change_names, axis=1)