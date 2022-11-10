# RM Fast Tracking System
This system is designed to enable planning team to check the availablity of raw materials for specific quantity of product. It is based on flask webpage and call sap to retrieve data and then upload the data to Google Sheets. 

Functions in app3.py:
1. index(): Redirect to login page of the ECM Fast Tracking System
2. login(): Validate the login credentials and direct to next page if credentials are correct, notice users failed to login in if credentials are wrong.
3. test(): Fetch the query data and process the input as dataframe, then call on sapautologin program

Functions in sapautologin.py
1. sapautologin: Get the query sku, date, and quantity and submit to Google Sheets
2. bomdata_download,find_marcdata,find_mbewdata,find_prdata,find_mb52,find_me2n: Get the specifc data from SAP Tables

The address for RM Fast Tracking is https://docs.google.com/spreadsheets/d/1q6tLUMal9t1EafNgs5ZeI-623e7Wh-bS-Sv1RwRIPQg

Attention:
1. Json file is needed to run the code, the service account need to apply for access to the sheet, otherwise the code is not able to run.
2. Please put Json file, folders, together with the code documents (2 of them) in the same environment path of Python.exe.
3. After running the app3.py, the ip address will occur if execute successfully, then please copy the ip address and open it via Chrome browser or Internet Explorer.
