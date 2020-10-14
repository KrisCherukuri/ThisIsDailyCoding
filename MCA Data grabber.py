"""
Steps
1.Go to MCA website to read the file
2.Load the file and do the necessary manipulations
3.Save a copy of file in the drive and sent an email to krishna.cherukuri@icicibank.com
"""

import pandas as pd 
import openpyxl as op
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

driver = webdriver.Chrome("C:/Users/Kishore/chromedriver_win32/chromedriver.exe")
driver.set_page_load_timeout("30")
driver.delete_all_cookies()
driver.get("https:www.mca.gov.in/mcafoportal/companiesRegReport.do")
time.sleep(30)
driver.quit()

wb=op.load_workbook("C:\\Users\\Kishore\\Downloads\\CompReg_14NOVEMBER2019.xlsx")
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
#HSR
dfJPNa = Karnataka[Karnataka['REGISTERED_OFFICE_ADDRESS'].str.contains("JP Nagar")]
#Domlur
dfDomlur = Karnataka[Karnataka['REGISTERED_OFFICE_ADDRESS'].str.contains("Domlur")]
#Writing df to an excel sheet and PDF
dfKoramangala.to_excel("D:\\Personal\\Google Drive\MCA\\Koramangala.xlsx")
dfDomlur.to_excel("D:\\Personal\\Google Drive\MCA\\Domlur.xlsx")
dfHSR.to_excel("D:\\Personal\\Google Drive\MCA\\HSR.xlsx")
dfJPNa.to_excel("D:\\Personal\\Google Drive\MCA\\JPNagar.xlsx")


