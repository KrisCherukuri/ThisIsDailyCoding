import openpyxl
import pandas as pd
import pdfkit as pdf
wb=openpyxl.load_workbook("D:/CompReg_21JULY2019.xlsx")
wb.sheetnames
sheet=wb['Indian Companies Registered']
sheet.delete_rows(1)
df = pd.DataFrame(sheet.values)
df.drop(df.columns[[4,6,7,10]], axis=1, inplace=True)
df=df.rename(columns=df.iloc[0]).drop(df.index[0])
Karnataka=df[df.STATE=='Karnataka']
Karnataka.head()
#Koramangala
dfKoramangala = Karnataka[Karnataka['REGISTERED_OFFICE_ADDRESS'].str.contains("Koramangala")]
#HSR
dfHSR = Karnataka[Karnataka['REGISTERED_OFFICE_ADDRESS'].str.contains("HSR")]
#Domlur
dfDomlur = Karnataka[Karnataka['REGISTERED_OFFICE_ADDRESS'].str.contains("Domlur")]
#Writing df to an excel sheet and PDF
dfKoramangala.to_excel("D:\\Z\\Google Drive\\MCA\\Koramangala.xlsx")
dfDomlur.to_excel("D:\\Z\\Google Drive\\MCA\\Domlur.xlsx")
dfHSR.to_excel("D:\\Z\\Google Drive\\MCA\\HSR.xlsx")
