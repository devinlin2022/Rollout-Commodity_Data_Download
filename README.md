# AM-Abnormal-Activity-Record-Notification
This program is designed to enable Operation Efficiency team to check the daily production issue in IM&LAM sub factory. It is based on python and call sap to retrieve data and then upload the data to Google Sheets.

Functions in python_email.py:

getname(): Input the name of vbs script for SAP to run.

editvbs(vbs_name,file1_name,script_name,combined_name): Process the vbs script by adding the content of "template.txt" to vbs script from getname function, and result returns as a name of completed vbs script.

call(file1_name,combined_name): Call the SAP to run complated VBS script from editvbs function.

dataupload(files,target_path): Process the data and save the data to the local folder.

xlsxgs(key,target_path,fname): Upload the processed excel file from local folder to Google Sheets.

CloseSAP(): Kill the task of SAPlogon.exe

Sendemail(): Send the notification email with updated Google Sheets link to OE supervisor to review and analyze the result.


The address for AM Abnormal Record is [https://docs.google.com/spreadsheets/d/1qteKQf0pSYQpE6s3RFAvbt2-sqzK1bUFqp1rT7uTSUM]

Attention:

1.Json file is needed to run the code, the service account need to apply for access to the sheet, otherwise the code is not able to run.
2.Please put Json file, TEXT, together with the code document in the same environment path of Python.exe.

