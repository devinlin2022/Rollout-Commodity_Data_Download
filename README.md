# ECM Fast Tracking System
This system is designed to enable ECM team to check the availablity of raw materials for specific quantity of product. It is based on flask webpage and call sap to retrieve data
and then upload the data to Google Sheets. 

Functions in app3.py:
1. index(): Redirect to login page of the ECM Fast Tracking System
2. login(): Validate the login credentials and direct to next page if credentials are correct, notice users failed to login in if credentials are wrong.
3. test(): Fetch the query data and process the input as dataframe, then call on sapautologin program

Functions in sapautologin.py
1. sapautologin: Get the query sku, date, and quantity and submit to Google Sheets
2. bomdata_download,find_marcdata,find_mbewdata,find_prdata,find_mb52,find_me2n: Get the specifc data from SAP Tables
