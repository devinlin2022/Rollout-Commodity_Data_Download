#pip install pymssql
#Part 1 Connect with SQL server
import pymssql
import pandas as pd
import numpy as np
import pygsheets

# conn = pymssql.connect(host='10.86.90.43', user='AS1\Devin lin', password='Abcd0816,', database='MESDB_ARCH')
conn = pymssql.connect(host='10.86.90.50', user='AS1\Devin lin', password='Abcd0816,', database='MESDB')
# conn = pymssql.connect(host='10.86.90.130', port='1433', user='AS1\Devin lin', password='Abcd0816,', database='DataMonitoring')
cursor = conn.cursor()

#Part 2 Get data from server
cursor.execute(f"SELECT [OBJECTID],[SUBOBJECTID],[MATNR],[CREDATETIME],[PACK_QTY] FROM [MESDB].[dbo].[NGSFR_MII_MGR_BAPI_HU_CREATE] Where MATNR like 'P%' and DateDiff(dd,CREDATETIME,getdate())<=7 order by CREDATETIME desc")
result = cursor.fetchall()
for i in result:
    print(i)
df = pd.DataFrame(result)
df.columns = ['HU Number', 'Batch Number', 'Material Number', 'Date', 'Pack Quantity']
df['HU Number'] = '00' + df['HU Number'].astype(str)

#Part 3 Process the data
gc = pygsheets.authorize(service_account_file=r'C:\Users\Devin Lin\Desktop\Project\civic-badge-353905-83ac72952c23.json')
#survey_url='https://docs.google.com/spreadsheets/d/1yWDaYwT3YSQGY2p7ScA8_KdfIw7EL4Jpoq9tZEQd6nY/edit#gid=0'
#wb_url = gc.open_by_url(survey_url)
key = '1O34ROoCS6Iy_Wh8c8900jo-dDrSIInuqEWGo6UTEwao'
wb_key = gc.open_by_key(key)
sh = wb_key.sheet1
df = pd.DataFrame(df)
sh.clear()
sh.set_dataframe(df,(1,1))
print("execute successfully")


