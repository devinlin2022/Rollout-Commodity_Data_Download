#pip install pymssql
#Part 1 Connect with SQL server
import pymssql
import pandas as pd
import numpy as np
# conn = pymssql.connect(host='10.86.90.43', user='AS1\Devin lin', password='Abcd0816,', database='MESDB_ARCH')
conn = pymssql.connect(host='10.86.90.50', user='AS1\Devin lin', password='Abcd0816,', database='MESDB')
# conn = pymssql.connect(host='10.86.90.130', port='1433', user='AS1\Devin lin', password='Abcd0816,', database='DataMonitoring')
cursor = conn.cursor()
#
#Part 2 Get data from server
cursor.execute(f"SELECT TOP 10 [OBJECTID],[SUBOBJECTID],[MATNR],[CREDATETIME] FROM [MESDB].[dbo].[NGSFR_MII_MGR_BAPI_HU_CREATE] Where MATNR like 'P%' order by CREDATETIME desc")
result = cursor.fetchall()
for i in result:
    print(i)
df = pd.DataFrame(result, columns=['Line', 'Shift_Desc', 'SKU', 'Shift_Start_Local', 'Order', 'Material', 'Description', 'Batch', 'Label', 'Quantity', 'Vendor', 'Verified By', 'BOM Status', 'Entered Time', 'Qty_Theoric'])
df.to_csv(r'C:\Users\bp_devin lin\Desktop\Project\rawdata.csv')

#Part 3 Process the data
df = pd.read_csv(r'C:\Users\bp_devin lin\Desktop\Project\rawdata.csv')
new=df.groupby(['Shift_Start_Local','Line','Shift_Desc','SKU','Material','Description','Qty_Theoric'])['Quantity'].agg('sum').reset_index()
new=df.groupby(['Shift_Start_Local','Line','Shift_Desc','SKU','Material','Description','Qty_Theoric'])['Quantity'].agg('sum').reset_index()
new=new.rename(columns={'Shift_Start_Local':'Date','SKU':'SKU#', 'Shift_Desc':'shift','Material':'P#','Qty_Theoric':'成品产出','Quantity':'包材总数'})
new['产出率']=new['成品产出']/(new['包材总数'])
new['损耗率']=(new['包材总数']-new['成品产出'])/new['成品产出']
new['产出率'] = new['产出率'].astype('str')
new['损耗率'] = new['损耗率'].astype('str')
a=new.loc[new['产出率'] >= "1.05"]
b=new.loc[new['产出率'] <= "0.85"]
c=pd.merge(a, b, how='outer', on=None, left_on=None, right_on=None)
c['产出率'] = c['产出率'].astype('float')
c['损耗率'] = c['损耗率'].astype('float')
c['产出率']=c['产出率'].apply(lambda x: format(x, '.2%'))
c['损耗率']=c['损耗率'].apply(lambda x: format(x, '.2%'))
order = ['Date', 'Line', 'shift', 'SKU#', 'P#', 'Description','包材总数', '成品产出', '损耗率', '产出率']
c=c[order]
c.to_csv(r'C:\Users\bp_devin lin\Desktop\Project\result.csv',encoding="utf_8_sig")

