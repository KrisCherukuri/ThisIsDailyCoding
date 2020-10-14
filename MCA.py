#1 Load data in the dataframe
import pandas as pd


file_name="C:/Users/Kishore/Desktop/2019 Acco help - All.xlsx"
sheetname= "Sheet2"

df = pd.read_excel(file_name,sheetname)
head = list(df)
df['Primary Phone no'] = '91'+ df['Primary Phone no'].astype(str)
for n in range(0,len(df.index)):
    df.URL[n]= str("https:/wa.me/" + df['Primary Phone no'] + str("?text=How%20are%20you%20?")
#2 Run a selenium webdriver in a loop where it sends all the messages
