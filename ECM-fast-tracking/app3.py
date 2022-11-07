import json
from flask import request, Flask, render_template, send_from_directory,jsonify
import pandas as pd
import sys, win32com.client
import win32api, win32gui, win32con, win32ui, time, os, subprocess
import time
import pandas as pd
import pygsheets
import clipboard
import os
import shutil
import pygsheets
import pandas as pd
# 创建程序
# web应用程序


app = Flask(__name__)

@app.route('/')
def index():
    return render_template("form.html")

@app.route('/login',methods = ['POST','GET'])
def login():
    username = request.form.get('username')
    password = request.form.get('pw')
    if username =='Alison' and password=='123':
        return render_template("index.html")
    else:
        return render_template('form.html',msg = "登录失败！")

@app.route('/running', methods=['POST', 'GET'])
def run():
    return render_template('index02.html',msg = "正在加载，请稍等20分钟")

@app.route('/test',methods = ['POST'])
def test():
    output = request.get_json()
    print(output)
    result = json.loads(str(output).replace("'", "\""))
    print(result)
    date = result.get('dt')
    date = date.replace('-', '/')
    sku = result.get('sku')
    qty = result.get('qty')
    dt_series = pd.Series(date, name='Date')
    sku_series = pd.Series(sku, name='SKU')
    qty_series = pd.Series(qty, name='qty')
    result_dataframe = pd.concat([dt_series,sku_series,qty_series],axis=1)
    result_dataframe.to_csv('./result.csv',index=False)
    subprocess.call(['cmd', '/c', 'sapautologin.py'])
    return render_template('index.html', msg="写入成功")
    # scriptdir = os.path.dirname(__file__)
    # batchfile = os.path.join(scriptdir, 'sapautologin.py')  # 请把.vbs修改为想要的名字,vbs文件需要和代码放在python内置环境下
    # subprocess.call(['cmd', '/c', os.path.realpath(batchfile), date])

app.run(debug=True,port=8002)


# class study01():
#     date = test()
#     print(date)oh