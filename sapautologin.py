import sys, win32com.client
import win32api, win32gui, win32con, win32ui, time, os, subprocess
import time
import pandas as pd
import pygsheets
import clipboard
import os
import shutil

# os.path.join('C:\Users\hayden tang\Documents\SAP\SAP GUI','M43415.txt')
def getquery():
    df = pd.read_csv('result.csv')
    query_date = df.loc[0, "Date"]
    query_sku = df.loc[0, "SKU"]
    query_qty = df.loc[0, "qty"]
    return query_sku,query_qty,query_date

def sap_login():
    sap_app = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"  # 您的saplogon程序本地完整路径
    subprocess.Popen(sap_app)
    time.sleep(1)
    flt = 0
    while flt == 0:
        try:
            hwnd = win32gui.FindWindow(None, "SAP Logon 760")
            flt = win32gui.FindWindowEx(hwnd, None, "Edit", None)  # capture handle of filter
        except:
            time.sleep(0.5)
    win32gui.SendMessage(flt, win32con.WM_SETTEXT, None, "[Production Users] CAP")
    win32gui.SendMessage(flt, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
    win32gui.SendMessage(flt, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)
    time.sleep(0.1)
    dlg = win32gui.FindWindowEx(hwnd, None, "Button", None)  # 登陆（0）
    win32gui.SendMessage(dlg, win32con.WM_LBUTTONDOWN, 0)
    win32gui.SendMessage(dlg, win32con.WM_LBUTTONUP, 0)
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return
    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return
    time.sleep(5)
    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return
    time.sleep(2)
    flag = 0
    while flag == 0:
        try:
            session = connection.Children(0)
            flag = 1
        except:
            time.sleep(0.5)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return
    time.sleep(1)
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "CNSTZ03" # 此次放入您的SAP登陆用户名
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "uS4siJsA@2" # 此次放入您的SAP登陆密码
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)
    try:

        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

    return session



def bomdata_download(session,sku):
    session.findById("wnd[0]").maximize()
    time.sleep(5)
    file_path=os.path.join(r'C:\Users\hayden tang\Documents\SAP\SAP GUI',r'{}bomdata.txt'.format(sku))
    if os.path.exists(file_path):
        os.remove(file_path)
    else:
        pass
    session.findById("wnd[0]/tbar[0]/okcd").text = "zmnu"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/mbar/menu[1]/menu[2]").select()
    session.findById("wnd[0]/usr/tblZMMZMTESTABCONTROL").getAbsoluteRow(2).selected = -1
    session.findById("wnd[0]/usr/tblZMMZMTESTABCONTROL/txtTAB-RPDES[0,2]").setFocus()
    session.findById("wnd[0]/usr/tblZMMZMTESTABCONTROL/txtTAB-RPDES[0,2]").caretPosition = 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = sku
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "CN12"
    session.findById("wnd[0]/usr/ctxtP_STLAN-LOW").text = "1"
    session.findById("wnd[0]/usr/txtP_FILE").setFocus()
    session.findById("wnd[0]/usr/txtP_FILE").caretPosition = 3
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/txtP_FILE").text = "C:\\DF"
    session.findById("wnd[0]/usr/txtP_FILE").setFocus()
    session.findById("wnd[0]/usr/txtP_FILE").caretPosition = 5
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById('wnd[1]/usr/ctxtDY_FILENAME').text='{}bomdata.txt'.format(sku)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/usr").verticalScrollbar.position = 1







    write_in_data=pd.read_table(file_path,encoding='unicode_escape',header=None)
    gc = pygsheets.authorize(
        service_file='C:\\Users\\hayden tang\\PycharmProjects\\cream loss\\.credentials\\client_secret.json')
    key = '1KWY5obxFiU2UmWLyV20GrxpNWvoG-iPCFKNKtnLI85w'
    sht = gc.open_by_key(key)  # 按照表格的key进行表格的读取
    raw_data = sht.worksheet_by_title(r'BOM')  # 选取指定的表单，对于表单进行更新
    raw_data.clear()
    raw_data.set_dataframe(write_in_data, (1, 1), copy_index=False, copy_head=False,nan=' ')  # 从第二行第二列开始更新为数据框test——table中的内容
    bom_list=write_in_data.iloc[:,9].values


    return session,bom_list


###  marc data download
def find_marcdata(session,bom_l,sku_l):
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zse16"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "MARC"
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)
    # copy multiple material and copy
    bom_lines = '\r\n'.join(bom_l)
    clipboard.copy(bom_lines)
    session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]").close()
    ### define factory cn12 and output it to txt file
    session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/usr/ctxtI2-LOW").text = "CN12"
    session.findById("wnd[0]/usr/ctxtI2-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtI2-LOW").caretPosition = 4
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[33]").press()


    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(52, "DEFAULT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 49
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "52"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    # session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"    #修改layout
    # session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
    # session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
    # session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()

    #### 这一步开始要改动一下

    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{}marc.txt".format(sku_l)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len("{}marc.txt".format(sku_l))
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

## open txtfile and upload it on google sheets
    file_path = os.path.join(r'C:\Users\hayden tang\Documents\SAP\SAP GUI', r"{}marc.txt".format(sku_l))

    write_in_data = pd.read_table(file_path, encoding='unicode_escape',skiprows=3,header=0,index_col=0)
    write_first = write_in_data.iloc[:,:18]
    write_second = pd.DataFrame(write_in_data['MS'])
    gc = pygsheets.authorize(
        service_file='C:\\Users\\hayden tang\\PycharmProjects\\cream loss\\.credentials\\client_secret.json')
    key = '1KWY5obxFiU2UmWLyV20GrxpNWvoG-iPCFKNKtnLI85w'
    sht = gc.open_by_key(key)  # 按照表格的key进行表格的读取
    raw_data = sht.worksheet_by_title(r'RM MARC')  # 选取指定的表单，对于表单进行更新
    raw_data.clear(start=(2,1))

    raw_data.set_dataframe(write_first, (2, 1), copy_index=False,copy_head=False,nan=' ')
    raw_data.set_dataframe(write_second, (2, 21), copy_index=False, copy_head=False,nan=' ')


    return session

### download mbew data for this same sku
def find_mbewdata(session,bom_l,sku_l):
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "zse16"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "Mbew"
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)
    bom_lines = '\r\n'.join(bom_l)
    clipboard.copy(bom_lines)
    session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]").close()
    session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/usr/ctxtI2-LOW").text = "cn12"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

 ### 选layout
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(4, "TEXT")
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()




    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{}MEBW.txt".format(sku_l)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len("{}MEBW.txt".format(sku_l))
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    time.sleep(2)
    # session.findById("wnd[1]").sendVKey(0)

    file_path = os.path.join(r'C:\Users\hayden tang\Documents\SAP\SAP GUI', r"{}MEBW.txt".format(sku_l))

    write_in_data = pd.read_table(file_path, encoding='unicode_escape', skiprows=3, header=0, index_col=0)

    gc = pygsheets.authorize(
        service_file='C:\\Users\\hayden tang\\PycharmProjects\\cream loss\\.credentials\\client_secret.json')
    key = '1KWY5obxFiU2UmWLyV20GrxpNWvoG-iPCFKNKtnLI85w'
    sht = gc.open_by_key(key)  # 按照表格的key进行表格的读取
    raw_data = sht.worksheet_by_title(r'RM MBEW')  # 选取指定的表单，对于表单进行更新
    raw_data.clear(start=(2,1))
    raw_data.set_dataframe(write_in_data, (2, 1), copy_index=False, copy_head=False,nan=' ')  # 从第二行第二列开始更新为数据框test——table中的内容
    # bom_list = write_in_data.iloc[:, 9].values

    return session


### prdata辅助函数
def generate_last_two(x,q_date):
    if x[r'ReqDate']!='':
        r_date=pd.to_datetime(x[r'ReqDate'])
        if r_date<pd.to_datetime(q_date):
            return True
        else:
            return False
    else:
        return ''
def generate_last(x):
    if x['last_two']!='':
        if x['last_two']:
            return 'Y'
        else:
            return 'N'

### download pr data for this sku
def find_prdata(session,sku_l,input_date,bom_l):
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMNU"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/mbar/menu[1]/menu[0]").select()
    session.findById("wnd[0]/tbar[0]/btn[83]").press()
    session.findById("wnd[0]/tbar[0]/btn[81]").press()
    session.findById("wnd[0]/tbar[0]/btn[81]").press()
    session.findById("wnd[0]/tbar[0]/btn[81]").press()
    session.findById("wnd[0]/tbar[0]/btn[81]").press()
    session.findById("wnd[0]/tbar[0]/btn[81]").press()
    session.findById("wnd[0]/tbar[0]/btn[81]").press()
    session.findById("wnd[0]/tbar[0]/btn[81]").press()
    time.sleep(10)
    session.findById("wnd[0]/usr/tblZMMZMTESTABCONTROL").getAbsoluteRow(44).selected = -1
    session.findById("wnd[0]/usr/tblZMMZMTESTABCONTROL/txtTAB-RPDES[0,7]").setFocus()
    session.findById("wnd[0]/usr/tblZMMZMTESTABCONTROL/txtTAB-RPDES[0,7]").caretPosition = 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = "CN12"
    session.findById("wnd[0]/usr/ctxtP_DATUM").setFocus()
    session.findById("wnd[0]/usr/ctxtP_DATUM").caretPosition = 5

    startdate = input_date
    enddate = pd.to_datetime(startdate) + pd.DateOffset(days=180)
    low_start_date=pd.to_datetime(startdate) - pd.DateOffset(days=5)
    enddate = enddate.strftime('%m/%d/%Y')
    low_start_date=low_start_date.strftime('%m/%d/%Y')

    session.findById("wnd[0]/usr/ctxtP_DATUM").text = enddate    #此处的日期为输入的日期往前推189天
    session.findById("wnd[0]/usr/ctxtP_DATUM").caretPosition = 10
    session.findById("wnd[0]/usr/btn%_S_DISPO_%_APP_%-VALU_PUSH").press()
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "DCI"
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "DCL"
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "ILI"
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus()
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 3
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]").close()
    session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()

    bom_lines = '\r\n'.join(bom_l)
    clipboard.copy(bom_lines)
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]").close()
    session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()

    session.findById("wnd[0]/usr/chkCB_PLARD").selected = -1
    session.findById("wnd[0]/usr/chkCB_IAR").selected = -1
    session.findById("wnd[0]/usr/chkCB_IAR").setFocus()
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/usr/radRB_RES").select()
    session.findById("wnd[0]/usr/chkCB_PRD").selected = -1
    session.findById("wnd[0]/usr/ctxtP_FDATE").setFocus()
    session.findById("wnd[0]/usr/ctxtP_FDATE").caretPosition = 0
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/usr/ctxtP_FDATE").text = low_start_date
    session.findById("wnd[0]/usr/ctxtP_TDATE").text = enddate
    session.findById("wnd[0]/usr/ctxtP_TDATE").setFocus()
    session.findById("wnd[0]/usr/ctxtP_TDATE").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById(
        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{}PR.txt".format(sku_l)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len("{}PR.txt".format(sku_l))
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    # session.findById("wnd[1]").sendVKey(0)



    ### 对于取得的

    file_path = os.path.join(r'C:\Users\hayden tang\Documents\SAP\SAP GUI', r"{}PR.txt".format(sku_l))
    write_in_data = pd.read_table(file_path, encoding='unicode_escape', skiprows=3, header=0, index_col=0)
    # write_in_data = write_in_data.fillna(' ')
    columns_list=[]
    for i in write_in_data.columns:
        columns_list.append(i.replace(' ',''))
    write_in_data.columns=columns_list
    write_in_data['last_two'] = write_in_data.apply(lambda x: generate_last_two(x,input_date), axis=1)
    write_in_data['last_one'] = write_in_data.apply(lambda x: generate_last(x), axis=1)
    last_two_columns=write_in_data[['last_two','last_one']]


    write_in_material = pd.DataFrame(write_in_data['MaterialNo.'])
    write_in_md = pd.DataFrame(write_in_data[r'MaterialDescr.'])
    write_in_rd = pd.DataFrame(write_in_data[r'ReqDate'])
    write_in_nrq = pd.DataFrame(write_in_data[r'NetReq.Qty'])
    gc = pygsheets.authorize(
        service_file='C:\\Users\\hayden tang\\PycharmProjects\\cream loss\\.credentials\\client_secret.json')
    key = '1KWY5obxFiU2UmWLyV20GrxpNWvoG-iPCFKNKtnLI85w'
    sht = gc.open_by_key(key)  # 按照表格的key进行表格的读取
    raw_data = sht.worksheet_by_title(r'RM PR')  # 选取指定的表单，对于表单进行更新
    max_r = raw_data.rows
    raw_data.clear(start=(2, 1),end=(max_r,4))
    raw_data.set_dataframe(write_in_material, (2, 1), copy_index=False, copy_head=False,nan=' ')
    raw_data.set_dataframe(write_in_md, (2, 2), copy_index=False, copy_head=False,nan=' ')
    raw_data.set_dataframe(write_in_rd, (2, 3), copy_index=False, copy_head=False,nan=' ')
    raw_data.set_dataframe(write_in_nrq, (2, 4), copy_index=False, copy_head=False,nan=' ')
    raw_data.set_dataframe(last_two_columns,(2,5),copy_index=False,copy_head=False,nan=' ')




    return session



####download mb52 data
### mb52辅助函数
def generate_mb52_last(x):
    if x[r'SLoc']!='':
        total=x[r'Unrestricted']+x[r'Transit/Transf.']+x['InQualityInsp.']+x[r'Restricted-Use']
        return total
    else:
        return ''

def find_mb52(session,bom_l,sku_l):
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "MB52"
    session.findById("wnd[0]").sendVKey(0)
    bom_lines = '\r\n'.join(bom_l)
    clipboard.copy(bom_lines)
    session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 3
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]").close()
    session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "CN12"
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/ctxtP_VARI").text = "ECOMM_LAYOUT"  #这个layout会对于前面marc的layout有影响么
    session.findById("wnd[0]/usr/ctxtP_VARI").setFocus()
    session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 12
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


    #### extract as txt file
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{}mb52.txt".format(sku_l)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len("{}mb52.txt".format(sku_l))
    session.findById("wnd[1]/tbar[0]/btn[11]").press()


    ## 对于导出的数据进行读取
    file_path = os.path.join(r'C:\Users\hayden tang\Documents\SAP\SAP GUI', r"{}mb52.txt".format(sku_l))
    #
    write_in_data = pd.read_table(file_path, encoding='unicode_escape', skiprows=1, header=0, index_col=0)
    columns_list = []
    for i in write_in_data.columns:
        columns_list.append(i.replace(' ', ''))
    write_in_data.columns = columns_list
    write_in_data = write_in_data[write_in_data['Material'].notnull()]

    write_in_material = pd.DataFrame(write_in_data['Material'])
    write_in_data[[ 'Unrestricted', r'Transit/Transf.', 'InQualityInsp.', 'Restricted-Use', 'Blocked', 'Returns','StockinTransit']].fillna(0,inplace=True)
    write_in_data[['Unrestricted', r'Transit/Transf.', 'InQualityInsp.', 'Restricted-Use', 'Blocked', 'Returns','StockinTransit']]=write_in_data[['Unrestricted', r'Transit/Transf.', 'InQualityInsp.', 'Restricted-Use', 'Blocked', 'Returns',\
                   'StockinTransit']].applymap(lambda x: str(x).replace(',','')).astype('float')
    write_in_data['last_row'] = write_in_data.apply(lambda x: generate_mb52_last(x), axis=1)
    write_in_other = write_in_data[['SLoc','Unrestricted', r'Transit/Transf.', 'InQualityInsp.', 'Restricted-Use', 'Blocked','Returns', 'StockinTransit','Batch']].fillna('')
    write_in_last=pd.DataFrame(write_in_data['last_row'])

    gc = pygsheets.authorize(
        service_file='C:\\Users\\hayden tang\\PycharmProjects\\cream loss\\.credentials\\client_secret.json')
    key = '1KWY5obxFiU2UmWLyV20GrxpNWvoG-iPCFKNKtnLI85w'
    sht = gc.open_by_key(key)  # 按照表格的key进行表格的读取
    raw_data = sht.worksheet_by_title(r'MB52')  # 选取指定的表单，对于表单进行更新
    max_r = raw_data.rows
    raw_data.clear(start=(2, 1), end=(max_r, 12))
    raw_data.set_dataframe(write_in_material, (2, 1), copy_index=False, copy_head=False,nan=' ')
    raw_data.set_dataframe(write_in_other, (2, 3), copy_index=False, copy_head=False,nan=' ')
    raw_data.set_dataframe(write_in_last, (2, 13), copy_index=False, copy_head=False, nan=' ')

    return session


#me2n辅助函数
def generate_me2n_last_two(x,q_date):
    if x[r'Del.Date']!='':
        r_date=pd.to_datetime(x[r'Del.Date'])
        if r_date<pd.to_datetime(q_date):
            return True
        else:
            return False
    else:
        return ''


##me2n
def find_me2n(session,bom_l,sku_l,input_date):
    input_date=input_date
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ME2N"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtSELPA-LOW").text = "WE101"
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "CN12"
    # session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = "M06738"



    ###copy bom_list
    bom_lines = '\r\n'.join(bom_l)
    clipboard.copy(bom_lines)
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]").close()
    session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()

    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").caretPosition = 6
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]").sendVKey(23)
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "{}me2n.txt".format(sku_l)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len("{}me2n.txt".format(sku_l))
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
### 对于下载的数据进行编辑
    file_path = os.path.join(r'C:\Users\hayden tang\Documents\SAP\SAP GUI', r"{}me2n.txt".format(sku_l))
    #
    write_in_data = pd.read_table(file_path, encoding='unicode_escape', skiprows=1, header=0, index_col=0)
    columns_list=[]
    for i in write_in_data.columns:
        columns_list.append(i.replace(' ',''))
    write_in_data.columns=columns_list

    write_in_data['last_two'] = write_in_data.apply(lambda x: generate_me2n_last_two(x,input_date), axis=1)
    write_in_data['last_one'] = write_in_data.apply(lambda x: generate_last(x), axis=1)
    last_two_columns=write_in_data[['last_two','last_one']]



    gc = pygsheets.authorize(
        service_file='C:\\Users\\hayden tang\\PycharmProjects\\cream loss\\.credentials\\client_secret.json')
    key = '1KWY5obxFiU2UmWLyV20GrxpNWvoG-iPCFKNKtnLI85w'
    sht = gc.open_by_key(key)  # 按照表格的key进行表格的读取
    raw_data = sht.worksheet_by_title(r'ME2N')  # 选取指定的表单，对于表单进行更新
    max_r = raw_data.rows
    write_in_other = write_in_data[['Purch.Doc.', 'Purch.Req.', 'Material', 'Rel.Qty', 'Del.Date','Tobedel.', 'Vendor/supplyingplant', 'Quantity', 'Sched.Qty','Crcy', 'Received']]
    raw_data.clear(start=(2, 1), end=(max_r, 11))
    raw_data.set_dataframe(write_in_other, (2, 1), copy_index=False, copy_head=False,nan=' ')
    raw_data.set_dataframe(last_two_columns,(2,12),copy_index=False, copy_head=False,nan=' ')

    return session



###
if __name__=='__main__':
    # find_sku,find_qty,find_date=getquery()
    # #将输入的日期，输入的sku，以及输入的查询数量输入到RM fast Tracking 表中
    # g_c = pygsheets.authorize(
    #     service_file='C:\\Users\\hayden tang\\PycharmProjects\\cream loss\\.credentials\\client_secret.json')
    # outer_key = '1KWY5obxFiU2UmWLyV20GrxpNWvoG-iPCFKNKtnLI85w'
    # sht_outer = g_c.open_by_key(outer_key)  # 按照表格的key进行表格的读取
    # raw_data_outer = sht_outer.worksheet_by_title(r'RM Checking')  # 选取指定的表单，对于表单进行更新
    # raw_data_outer.update_value('B5',find_sku)
    # raw_data_outer.update_value('G4',find_date)
    # raw_data_outer.update_value('C5',find_qty)
    ### sap自动登陆
    new_session=sap_login()
    # #bom data 的抓取
    current_session,bomdata=bomdata_download(new_session,find_sku)
    # # marc数据的下载
    current_session=find_marcdata(current_session,bomdata,find_sku)
    # # mbew数据的下载
    current_session=find_mbewdata(current_session,bomdata,find_sku)
    # # #pr数据的下载
    current_session=find_prdata(current_session,find_sku,find_date,bomdata)
    # # # #mb52数据的下载
    current_session=find_mb52(current_session,bomdata,find_sku)
    # # #me2n数据的下载
    current_session=find_me2n(current_session,bomdata,find_sku,find_date)



