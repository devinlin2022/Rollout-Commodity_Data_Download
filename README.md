# AM-Abnormal-Activity-Notification
This program is designed to enable Operation Efficiency team to check the daily production issue in IM&LAM sub factory. It is based on python and call sap to retrieve data and then upload the data to Google Sheets.

Functions in python_email.py:

index(): Redirect to login page of the ECM Fast Tracking System
login(): Validate the login credentials and direct to next page if credentials are correct, notice users failed to login in if credentials are wrong.
test(): Fetch the query data and process the input as dataframe, then call on sapautologin program
Functions in sapautologin.py

sapautologin: Get the query sku, date, and quantity and submit to Google Sheets
bomdata_download,find_marcdata,find_mbewdata,find_prdata,find_mb52,find_me2n: Get the specifc data from SAP Tables
The address for RM Fast Tracking is https://docs.google.com/spreadsheets/d/1qteKQf0pSYQpE6s3RFAvbt2-sqzK1bUFqp1rT7uTSUM

Attention:

Json file is needed to run the code, the service account need to apply for access to the sheet, otherwise the code is not able to run.
Please put Json file, folders, together with the code documents (2 of them) in the same environment path of Python.exe.

