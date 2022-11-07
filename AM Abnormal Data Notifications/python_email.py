import os

# os.system('pip install openpyxl')
# os.system('pip install subprocess')
# os.system('pip install pandas')
# os.system('pip install datetime')

import os
import os
import subprocess
import glob
import pandas as pd
from datetime import datetime as dt
import pygsheets
import time
import datetime
import os
import smtplib
from email.mime.text import MIMEText
from email.header import Header
# import pygsheets

def getname():
    inputname="IW29_下载AM失效问题2"
    vbs_name = inputname + '.vbs'
    script_name = 'template.txt'
    file1_name = inputname + '_1.vbs'
    combined_name = inputname + '_2.vbs'
    return vbs_name,script_name,file1_name,combined_name

def editvbs(vbs_name,file1_name,script_name,combined_name):
    file1 = open(vbs_name, encoding='gb18030', errors='ignore')
    a=file1.read()
    set_print = False
    for line in a.split('\n'):
        if line.startswith('session.findById'):
            set_print = True
        if set_print:
            print(line)
            with  open(file1_name, "a") as file:
                file.write('\n'+line)
                file.close()

    with open(script_name, encoding='gb18030', errors='ignore') as fp:
        data = fp.read()

    with open(file1_name, encoding='gb18030', errors='ignore') as fp:
        data2 = fp.read()

    data += "\n"
    data += data2

    with open(combined_name, 'w', encoding='utf-8') as fp:
        fp.write(data)

    return file1_name, combined_name

def call(file1_name,combined_name):
    folder_path = r'C:\Users\devin lin\Documents\SAP\SAP GUI'
    file_type = r'\IW29.xls'
    files = folder_path + file_type
    print(files)
    if os.path.exists(files):
        os.remove(files)
        print('Delete successfully!')
    else:
        print('File not exist')


    target_path = r'O:\Shared drives\IM Shared Drive'#这里需要修改与google drive连接的地址，目前采用开发者设置默认地址
    key = '1qteKQf0pSYQpE6s3RFAvbt2-sqzK1bUFqp1rT7uTSUM'
    vbs_name = combined_name

    scriptdir = os.path.dirname(__file__)
    batchfile = os.path.join(scriptdir, combined_name)#请把.vbs修改为想要的名字,vbs文件需要和代码放在python内置环境下
    print(os.path.realpath(batchfile))
    subprocess.call(['cmd', '/c', os.path.realpath(batchfile)])

    os.remove(os.path.realpath(batchfile))
    batchfile2 = os.path.join(scriptdir, file1_name)
    os.remove(os.path.realpath(batchfile2))
    return files, target_path, key

def dataupload(files,target_path):
    # folder_path = folder_path
    # file_type = r'\*xls'
    # files = glob.glob(folder_path + file_type)
    # max_file = max(files, key=os.path.getctime)
    # print(max_file)
    df = pd.read_table(files, encoding="utf-16",skiprows=1)
    df.drop(columns=df.columns[0], axis=1, inplace=True)
    mask = '%m%d'
    dte = dt.now().strftime(mask)
    file_name = vbs_name.split('.')[0]
    fname = file_name + '_{}.csv'
    fname = fname.format(dte)
    target_path = target_path + "\\" + fname
    print(target_path )
    df.to_csv(target_path)
    print("Copy to share drive successfully!")
    return fname

def xlsxgs(key,target_path,fname):
    gc = pygsheets.authorize(service_account_file=r'C:\Users\devin lin\PycharmProjects\pythonProject\venv\client_secret.json')
    # survey_url='https://docs.google.com/spreadsheets/d/1yWDaYwT3YSQGY2p7ScA8_KdfIw7EL4Jpoq9tZEQd6nY/edit#gid=0'
    # wb_url = gc.open_by_url(survey_url)
    key = key
    wb_key = gc.open_by_key(key)
    sh = wb_key.worksheet_by_title(r'Sheet1')#Change to the sheet name you want
    open_file = target_path + "\\" + fname
    df = pd.read_csv(open_file)
    df = pd.DataFrame(df)
    df.drop(columns=df.columns[0], axis=1, inplace=True)
    df['Notif.date'] = pd.to_datetime(df['Notif.date']).dt.strftime('%Y-%m-%d')
    max_r = sh.rows
    sh.clear(start=(1, 1), end=(max_r, 11))
    sh.set_dataframe(df, (1, 1))
    print("Paste to Google Sheets successfully")
#    sh.update_values_batch(['A:K'], [1,2,3,4,5,6,7,8,9,10,11,12], 'COLUMNS')

def close_SAP():
    os.system('taskkill /im saplogon.exe /t /f')
    print("sap close successfully")


def sendemail():
    today = datetime.datetime.now().date()
    body = 'Hello, ' \
           'AM失效记录已更新,请查看' \
           'link: https://docs.google.com/spreadsheets/d/1qteKQf0pSYQpE6s3RFAvbt2-sqzK1bUFqp1rT7uTSUM/edit#gid=0'
    sender = 'from@colpal.com'
    receivers = ['devin_lin@colpal.com','rosa_liu@colpal.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱

    # 三个参数：第一个为文本内容，第二个 plain 设置文本格式，第三个 utf-8 设置编码
    message = MIMEText(body, 'html', 'utf-8')

    subject = 'AM失效记录' + today.strftime('%Y-%m-%d')
    message['Subject'] = Header(subject, 'utf-8')

    print(body)
    try:
        smtpObj = smtplib.SMTP('lnmailhost2.esc.win.colpal.com')
        smtpObj.sendmail(sender, receivers, message.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException:
        print("Error: 无法发送邮件")

vbs_name,script_name,file1_name,combined_name = getname()
file1_name, combined_name = editvbs(vbs_name,file1_name,script_name,combined_name)
files, target_path, key=call(file1_name,combined_name)
# vbs_name,folder_path,target_path,key=getquery(combined_name)
fname=dataupload(files,target_path)
xlsxgs(key,target_path,fname)
close_SAP()
sendemail()
