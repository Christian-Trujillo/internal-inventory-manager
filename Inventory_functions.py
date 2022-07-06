from cmath import nan
from matplotlib.pyplot import show
import openpyxl as xl
import os, re
from googleapiclient.discovery import build
from google.oauth2 import service_account
import PySimpleGUI as sg
import pandas as pd
import numpy as np
import json
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
from calendar import month, weekday
from datetime import datetime
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime,date, timedelta
from dateutil.relativedelta import relativedelta
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import shutil, re
from webdriver_manager.chrome import ChromeDriverManager
from threading import Thread
import pickle
import queue
import win32gui
signature='CT'
### define variables needed for other funcs ###
def Initialize(Queue):
    
    global SCOPES
    global SERVICE_ACCOUNT_FILE
    global creds
    global service
    global CONTAINERS_ID
    global TRANSFERS_ID
    global INV_SAFETY_ID
    global FORECAST_ID
    global IIMCHANGELOG_ID
    global path
    global sheet
    global sku_list
    global reduced_sku_list
    global sku_dict
    global sku_details
    global dates
    global period_to_weeks
    global daterange
    global forecast_list
    global today
    global containers
    global inv_safety
    global transfers
    global chrome_path
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    path = os.getcwd()
    # get info from keys.json
    SERVICE_ACCOUNT_FILE = path+r'\keys.JSON'
    creds= None
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes = SCOPES)
    #client id = '318827037112-eq655i5b2ns04g8pobdlbrch6g3nqu9u.apps.googleusercontent.com'
    service = build('sheets', 'v4', credentials=creds)
    # The ID and range of a sample spreadsheet.
    CONTAINERS_ID = '1d8hRkptQwV9VPLhgJLsVgYdwk8exz-GdF5oBdwjFbQE'
    TRANSFERS_ID = '1Di-2F9A4xSndZGOx7plq2JDzzConxMy0KRjeHO0a_BI'
    # TRANSFERS_ID = '180OiU7000XFu3iN6BL0Edj1iKwmTEgfJ26xbT3nzoqk'
    INV_SAFETY_ID = '1Li0W6bfRA-x80TyOt7yGus5y12AzMucom4NdrfWfqRU'
    FORECAST_ID = '13sJ9Iyp3rfeas7du6xwYeNLquORTE1OF75hQMKK7kdY'
    IIMCHANGELOG_ID = '1MX4WsjLRB9MWtqNM870-qD3cz66UQ67PstEZn6-Zc1o'
    # Call the Sheets API
    sheet = service.spreadsheets()
    ### forecast variables ###
    containers,inv_safety, transfers, change_log_s,change_log_i = read_sheets()
    
    with open('references.json','r+') as f:
        data=json.load(f) 
        sku_dict = data['sku_dict']
        reduced_sku_list = data['reduced_sku_list']
        sku_details = data['sku details']
        sku_list = data['sku_list']
    ### list of dates to use in sales data ###
    def daterange(start_date, end_date):
        for n in range(int((end_date - start_date).days)):
            yield start_date + timedelta(n)
    today=datetime.today().strftime("%y-%m-%d")
    period_to_weeks={'1 Week':1,'2 Weeks':2,'1 Month':4,'2 Months':8,'3 Months':13,'4 Months':17,'5 Months':22,'6 Months':26,'7 Months':31,'8 Months':35,'9 Months':39}
    forecast_list=None
    options=webdriver.ChromeOptions()
    options.headless = True
    chrome_path = ChromeDriverManager().install()
    SC_driver = webdriver.Chrome(executable_path=chrome_path,chrome_options=options)
    SC_driver.get(f"https://df.cwa.sellercloud.com/login.aspx?ReturnUrl=%2f")
    SC_login(SC_driver)

    Queue.put((SCOPES,SERVICE_ACCOUNT_FILE,service , CONTAINERS_ID ,TRANSFERS_ID,INV_SAFETY_ID,FORECAST_ID,IIMCHANGELOG_ID,path,sheet,sku_list,reduced_sku_list,sku_dict,sku_details,change_log_s,change_log_i,period_to_weeks,daterange,forecast_list,today,containers,inv_safety, transfers,SC_driver))
    return SCOPES,SERVICE_ACCOUNT_FILE,service , CONTAINERS_ID ,TRANSFERS_ID,INV_SAFETY_ID,FORECAST_ID,IIMCHANGELOG_ID,path,sheet,sku_list,reduced_sku_list,sku_dict,sku_details,change_log_s,change_log_i,period_to_weeks,daterange,forecast_list,today,containers,inv_safety, transfers,SC_driver
### pull live data from needed google sheets, return as a pandas DataFrame ###
def read_sheets():
    containers = pd.DataFrame(sheet.values().get(spreadsheetId=CONTAINERS_ID,range="'Current Containers'!A3:z100").execute().get('values',[]))
    inv_safety = pd.DataFrame(sheet.values().get(spreadsheetId=INV_SAFETY_ID,range="'inventory and safeties'!A2:d150").execute().get('values',[]))
    transfers = pd.DataFrame(sheet.values().get(spreadsheetId=TRANSFERS_ID,range="'Warehouse Transfers'!A2:T200",valueRenderOption = 'FORMULA').execute().get('values',[]))
    change_log_s = sheet.values().get(spreadsheetId=IIMCHANGELOG_ID,range="'wholesale safeties changes'!A1:e1000").execute().get('values',[])
    change_log_i = sheet.values().get(spreadsheetId=IIMCHANGELOG_ID,range="'inventory adjustments'!A1:f1000").execute().get('values',[])

    containers = containers.replace([None],['']).values.tolist()
    transfers = transfers.replace([None],['']).values.tolist()
    inv_safety = inv_safety.replace([None],['']).values.tolist()
    containers.remove(containers[0])
    transfers.remove(transfers[0])
    return containers,inv_safety, transfers,change_log_s,change_log_i
### enter str in format of sku1=qty1 / sku2=qty2/... . Returns dictionary of each container and their item qtys ###
def grab_qty(str):
    ### regex list for finding items in containers ###
    regexlist = []
    for item in sku_list:
        regexlist.append(rf'{item}=\d+')
        regexlist.append(rf'{item}= \d+')
        regexlist.append(rf'{item} =\d+')
        regexlist.append(rf'{item} = \d+')
    mydict={}
    xlist =re.findall(r"(?=("+'|'.join(regexlist)+r"))", str.upper())
    for i in range(len(xlist)):
        try: xlist[i]=xlist[i].split('=')
        except: xlist.remove(xlist[i])
    for item in xlist:
        mydict[item[0]]=item[1]
    return mydict
### enter sku as item, uses grab_qty func to return dict of containers that contain the specific sku, {container#1:qty1, container#2, qty2...}
### only returns containers that have not been recieved and added into SC ###
def search_containers(item):
    containers_with_item = {}
    for row in containers:
        if row[15]=='':
            item_qty = grab_qty(row[11])
            if item.upper() in item_qty:
                containers_with_item[row[0]]=item_qty[item.upper()]
    return containers_with_item
### inputs data from containers recieved at storage warehouses into transfers page in speific format ###
def update_transfer():
    containers_updated=[]
    B=0
    M=0
    O=0
    for i in range(len(transfers)):
        if transfers[i][0] == '':
            benson_len = i 
            break
    for i in range(len(transfers)):
        if transfers[i][7] == '':
            mag_len = i
            break
    for i in range(len(transfers)):
        if transfers[i][14] == '':
            ont_len = i
            break
    for i in range(len(containers)):
        item_qty = grab_qty(containers[i][11])
        if containers[i][14].upper() == 'X' and containers[i][12].upper().find('UPDATED') == -1 and  containers[i][12].find('BROOKS') == -1:
            for key in item_qty:
                if containers[i][12].upper().find('BENSON') != -1:
                    transfers[benson_len+B][0] = str(containers[i][9])[:8]
                    transfers[benson_len+B][1] = key
                    transfers[benson_len+B][2] = '-'+item_qty[key]
                    transfers[benson_len+B][3] = containers[i][0]
                    containers[i][12] = 'BENSON - UPDATED'
                    containers[i][17]= ''
                    containers_updated.append(containers[i][0])
                    B+=1
                if containers[i][12].upper().find('MAGNOLIA') != -1:
                    transfers[mag_len+M][7] = str(containers[i][9])[:8]
                    transfers[mag_len+M][8] = key
                    transfers[mag_len+M][9] = '-'+item_qty[key]
                    transfers[mag_len+M][10] = containers[i][0]
                    containers[i][12] = 'MAGNOLIA - UPDATED'
                    containers[i][17]= ''
                    containers_updated.append(containers[i][0])
                    M+=1
                if containers[i][12].upper().find('ONTARIO') != -1:
                    transfers[ont_len+O][14] = str(containers[i][9])[:8]
                    transfers[ont_len+O][15] = key
                    transfers[ont_len+O][16] = '-'+item_qty[key]
                    transfers[ont_len+O][17] = containers[i][0]
                    containers[i][12] = 'ONTARIO - UPDATED'
                    containers[i][17]= ''
                    containers_updated.append(containers[i][0])
                    O+=1
    sheet.values().update(spreadsheetId=TRANSFERS_ID, range="'Warehouse Transfers'!A3", valueInputOption='USER_ENTERED', body={'values':transfers}).execute()
    sheet.values().update(spreadsheetId=CONTAINERS_ID, range="'Current Containers'!A4", valueInputOption='USER_ENTERED', body={'values':containers}).execute()
    return containers_updated
### uses regex to find most recent exports for LOW INV REPORT and Safety export, if none are manually selected ###
def find_general_inv_files(inv, sfty):
    low_inv_wkbk = xl.load_workbook(inv)
    low_inv_obj=pd.DataFrame(low_inv_wkbk.active.values)
    sfty_wkbk = xl.load_workbook(sfty)
    sfty_obj=pd.DataFrame(sfty_wkbk.active.values)

    return low_inv_obj,sfty_obj
### inputs data from inventory files from find_general_inv_files into general inventory google sheets ###
def update_inv_safety(low_inv_obj, sfty_obj):
    ### create dictionary for safety/min qty per sku ###
    sheet = service.spreadsheets()
    safety_list = sheet.values().get(spreadsheetId=INV_SAFETY_ID,range="'inventory and safeties'!i2:k66").execute().get('values',[])
    safety_dict={}
    for item in safety_list:
        safety_dict[item[0]]=[int(item[1]),int(item[2])]
    
    ### current inventories and safeties as df###
    low_inv = low_inv_obj
    sfty=sfty_obj
    ### 
    inv = pd.DataFrame(sheet.values().get(spreadsheetId=INV_SAFETY_ID,range="'inventory and safeties'!a2:f111").execute().get('values',[]))
    past_inv=inv.filter([0,2]).astype({2:int})
    inv = inv.filter([0,5])
    index = pd.DataFrame(sku_list)
    sfty = sfty.filter([63,64])
    sfty[63][sfty[63]=='COS-640SLTX-E']='COS-640STX-E'
    sfty.columns=[0,1]
    sfty=sfty[1:]
    sfty[1]=sfty[1].astype(int)
    low_inv = low_inv[1:]
    low_inv = low_inv.astype({2:int, 3:int})

    result = pd.merge(index, low_inv.filter([0,2,3]), on=0)
    result = pd.merge(result, sfty, how="left", on=0)
    result[5]=0
    result[6]=''
    result = pd.merge(result, inv, how="left", on=0)
    result = pd.merge(result, past_inv, how="left", on=0)
    result.columns=['SKU', 'AGGREGATE','PHYSICAL','SAFETY', 'ON WATER', 'ERRORS','NOTES','past inv']

    for sku in sku_list:
        if sku not in safety_dict.keys():
            safety_dict[sku]=[0,0]
    safety_dict=pd.DataFrame(safety_dict,index=['safety','min']).transpose()
    result.index=result['SKU']
    # past_inv.index=past_inv[0]
    # past_inv.sort_index()
    safety_dict=safety_dict.loc[result.index.tolist(),:]
    result = result.replace([nan],[0])
    ### on water ###
    result['ON WATER']=item_quantity()

    ### fill errors col ###   
    conditions = [(result["AGGREGATE"]<1 ) & (result['SAFETY']>safety_dict['min']),
    (result["PHYSICAL"]==0) | ((result["AGGREGATE"]<1) & (result['SAFETY'] <= safety_dict['min'])),
    ((result["AGGREGATE"]>safety_dict['safety']) & (result['SAFETY']<safety_dict['safety'])) |
    ((result["PHYSICAL"]>result['past inv']) & (result['SAFETY']<safety_dict['safety'])),
    result["AGGREGATE"]<0,
    result['SAFETY']<safety_dict['min'],
    result['SAFETY']<safety_dict['safety']]

    choices = ['Lower Safety; \n Item is not selling',
    'Item OOS',
    'Item may be back in stock;\n please Cycle Count',
    'Negative Aggregate;\n check backorders and Cycle Count',
    'Minimum Reduced','Safety Reduced']
    
    result['ERRORS'] = np.select(conditions, choices, default='')
    result.drop(['past inv'],1, inplace=True)
    result["NOTES"].replace([0,'0'],'', inplace=True)
    result = result.replace([None],[''])
    result = result.replace([nan],['']) 
    result['SAFETY']=result['SAFETY'].replace([''],[0])
    result=[result.columns.tolist()] + result.values.tolist()

    sheet.values().update(spreadsheetId=INV_SAFETY_ID, range=f'inventory and safeties!G1', valueInputOption='USER_ENTERED', body={'values':[[datetime.today().strftime('%B %e')]]}).execute()
    sheet.values().update(spreadsheetId=INV_SAFETY_ID, range=f'inventory and safeties!a1', valueInputOption='USER_ENTERED', body={'values':result}).execute()
### back in stock sheet ###
def update_back_in_stock(low_inv_obj):
    sheet = service.spreadsheets()
    df = pd.DataFrame(sheet.values().get(spreadsheetId=INV_SAFETY_ID,range="'inventory and safeties'!A2:C200").execute().get('values',[]))
    df.index=df[0]; df=df.drop([0,2],axis=1);df.astype(int)
    low_inv_obj = low_inv_obj.drop(0);low_inv_obj.index=low_inv_obj[0]; low_inv_obj=low_inv_obj.drop([0,1,3],axis=1); 
    df=df.merge(low_inv_obj, left_on=df.index, right_on=low_inv_obj.index,how='left')
    df.columns=['SKU', 'Yesterday','Today']
    df.index=df['SKU']; df=df.drop('SKU',axis=1);df=df.astype({'Yesterday':int,'Today':int})
    df['B.I.S']="" ; df['B.I.S'][df['Yesterday']<5] = np.where(df['Today']>df['Yesterday'], 'X', "")
    if not df['Today'].equals(df['Yesterday']):
        df_list=[df.columns.tolist()] + df.reset_index().values.tolist()
        df_list[0].insert(0,'SKU')
        sheet.values().update(spreadsheetId=INV_SAFETY_ID, range=f"'back in stock (temp)'!a1", valueInputOption='USER_ENTERED', body={'values':df_list}).execute()

### editing and adding of containers into containers sheets ###
def update_containers(container_number, freight_forwarder,ETA, contents, MNFCR, Notes_1,Notes_2):
    
    container_numbers = {}
    for i in range(len(containers)):
        container_numbers[containers[i][0]]=i
    if container_number not in container_numbers:
        containers.append([container_number, freight_forwarder, '','',ETA,'','','','','','',contents,'',MNFCR,'','','',Notes_1,Notes_2,'','','','','',''])
        sheet.values().update(spreadsheetId=CONTAINERS_ID, range="'Current Containers'!A4", valueInputOption='USER_ENTERED', body={'values':containers}).execute()

    elif container_number in container_numbers:
        popup = sg.Window('Overwrite Container Log?',[[sg.Text('Container is already in log')],[sg.Text('Overwrite With Entered Values?')],[sg.Button('yes'),sg.Button('no')]])
        event,values = popup.read()
        if event =='yes':
            row = containers[container_numbers[container_number]]
            new_row = [container_number, freight_forwarder, '','',ETA,'','','','','','',contents,'',MNFCR,'','','',Notes_1,Notes_2,'','','','','','']
            for i in range(len(new_row)):
                if new_row[i]=='' and row[i]!='':
                    new_row[i]=row[i]
            containers[container_numbers[container_number]] = new_row
            popup.close()
            sheet.values().update(spreadsheetId=CONTAINERS_ID, range="'Current Containers'!A4", valueInputOption='USER_ENTERED', body={'values':containers}).execute()

        if event =='no':
            popup.close()
### returns DataFrame of totals for low_inv_report AGG + on water inv (from search_containers) ###
def item_quantity(add_agg = False):
    '''Searches Container Logs using Regex to find and sum quantaties of items on water\n
    Returns Pandas Dataframe indexed by sku \n\n 
    add_agg -> adds current aggregate inventory to on water quantities, useful for determining how much will be in stock in ~1 month \n\n
    '''
    item_qty={}
    for row in inv_safety:
        if add_agg:
            item_qty[row[0]]=int(row[1])
        else:
            item_qty[row[0]]=0
    for item in sku_list:
        total=0
        if search_containers(item)!={}:
            dict = search_containers(item)
            for qty in dict.values():
                total+=int(qty)
            item_qty[item]+=total
    df= pd.DataFrame(item_qty, index=['Stock']).transpose()
    return df
### downloads sales export and updates directory ###
def SC_login(driver):
    email = driver.find_element(By.NAME , "ctl00$ContentPlaceHolder1$txtEmail")
    pwd = driver.find_element(By.NAME , "ctl00$ContentPlaceHolder1$txtPwd")
    email.clear()
    pwd.clear()
    email.send_keys("trujillochristian.cosmo@gmail.com")
    pwd.send_keys("Quertzy062813!"+Keys.ENTER)
### downoad SC exports needed for updating sheets ###
def download_sales(window):
    window.write_event_value('--salesupdate--', 'Updating Files')
    ### opens Chrome ###
    options=webdriver.ChromeOptions()
    options.headless = True

    prefs={"download.default_directory":os.getcwd()+r'\exports'}
    options.add_experimental_option("prefs",prefs)
    driver = webdriver.Chrome(executable_path=chrome_path,chrome_options=options)  
    ### looks up value in Chrome ###
    driver.get("https://df.cwa.sellercloud.com/DashboardV2/Reports/ReportV2_ProductQtySoldByDay.aspx")
    SC_login(driver)
    ### waits until title contains text ###
    WebDriverWait(driver, 30).until(EC.title_contains('Qty Sold By Product Per Day Report')) #This is a dummy element
    ### checks if "text" in </title> ###
    assert "Qty Sold By Product Per Day Report - SellerCloud" == driver.title
    driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$txtFromDate").send_keys(start_date)
    driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$txtToDate").send_keys(end_date)
    run_report=driver.find_element(By.ID, "ContentPlaceHolder1_btnRunReport")
    run_report.click()
    # run_report.get_property()
        
    # webdriver.ActionChains(driver).click(run_report).perform()
    # WebDriverWait(driver,180).until(EC.element_to_be_selected ((By.NAME, 'ctl00$ContentPlaceHolder1$imgExcel')))
    export_excel=driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$imgExcel')
    export_excel.click()
    waiting_for_dl=True
    while waiting_for_dl:
        try:
            shutil.move(path +r'\exports\ProductQuantitySoldByDay.xlsx' , path + rf'\exports\Sales Data.xlsx')
            break
        except: sleep(1)
def download_safeties(window,queue):
    
    try:
        # window.write_event_value('--invupdate--', 'Updating Files')
        options=webdriver.ChromeOptions()
        # options.headless = True
        prefs={"download.default_directory":os.getcwd()+r'\exports'}
        options.add_experimental_option("prefs",prefs)
        driver = webdriver.Chrome(executable_path=chrome_path,chrome_options=options)    
        driver.get("https://df.cwa.sellercloud.com/Orders/Orders_details.aspx?ID=7808146")
        SC_login(driver)
        WebDriverWait(driver, 30).until(EC.title_contains('Order 7808146')) 
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ActionList").send_keys('Export Order')
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ImageButton1").click() 
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ddlFileType").send_keys('Excel')
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnExportOrders").click() 
        regex = re.search(r'JobID=\d+',driver.page_source)
        JobId = regex.group()[6:]
        today=datetime.today().strftime("%m/%d/%Y")
        driver.get(f'https://df.cwa.sellercloud.com/MyAccount/QueuedJobs.aspx?UserID=2691946&JobType=-1&SubmittedOnStartDate={today}&SubmittedOnEndDate={today}&Status=-1')
        waiting_for_dl=True
        while waiting_for_dl:
            try:
                driver.find_element(By.ID,'ContentPlaceHolder1_QueuedJobsList_grdMain_ctl00_ctl04_btnViewOutput').click()
                break
            except:
                sleep(10)
                driver.refresh()
        while waiting_for_dl:
            try:
                shutil.move(path +rf'\exports\Orders_Export_{JobId}.xlsx' , path + r'\exports\Safeties.xlsx')
                break
            except: sleep(1)
    except: queue.put('--Failure--')
    driver.close()
def download_inv(window,queue):
    try:
        # window.write_event_value('--invupdate--', 'Updating Files')
        mod_sku_list = ''
        for sku in sku_list[:-1]:
            mod_sku_list+=sku+' , '
        mod_sku_list+=sku_list[-1]
        options=webdriver.ChromeOptions()
        # options.headless = True
        prefs={"download.default_directory":os.getcwd()+r'\exports'}
        options.add_experimental_option("prefs",prefs)
        driver = webdriver.Chrome(executable_path=chrome_path,chrome_options=options)
        driver.get("https://df.cwa.sellercloud.com/Inventory/ManageInventory.aspx?CompanyIDList=&sku=&active=1&rowsperPage=50&inventoryFrom=-2147483648&inventoryTo=-2147483648&SavedSearchName=&SKUUseWildCards=False&OrderID=0&InventoryViewMode=0&InventoryQtyFilterMode=0&SortBy=bvc_Product.ID&SortByDirection=ASC&")
        SC_login(driver)
        sleep(2)
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$txtSKU").send_keys(mod_sku_list)
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnSearchNow").click()
        driver.find_element(By.ID, "ContentPlaceHolder1_chkSelectAll").click()
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnExportProducts").click()
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ddlExportformat").send_keys('Excel')
        driver.find_element(By.XPATH,"//a[@href='/inventory/CustomExport.aspx']").click()
        driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ddlTemplate").send_keys('AVC AGG/PHYS')
        driver.find_element(By.ID, "ContentPlaceHolder1_btnLoadTemplate").click()
        driver.find_element(By.ID, "ContentPlaceHolder1_ddlExportFileFormat").click()
        sleep(1)
        driver.find_element(By.NAME,"ctl00$ContentPlaceHolder1$btnExport").click()
        sleep(3)
        regex = re.search(r'JobID=\d+',driver.page_source)
        JobId = regex.group()[6:]
        driver.get(f'https://df.cwa.sellercloud.com/MyAccount/QueuedJobs.aspx?UserID=2691946&JobType=-1&SubmittedOnStartDate={today}&SubmittedOnEndDate={today}&Status=-1')
        waiting_for_dl=True
        while waiting_for_dl:
            try:
                driver.find_element(By.ID,'ContentPlaceHolder1_QueuedJobsList_grdMain_ctl00_ctl04_btnViewOutput').click()
                break
            except:
                sleep(10)
                driver.refresh()
        while waiting_for_dl:
            try:
                shutil.move(path +rf'\exports\\{JobId}.xlsx' , path + r'\exports\Inventory.xlsx')
                break
            except: sleep(1)
    except: 
        queue.put('--Failure--')
    driver.close()
def download_vel(window):
    window.write_event_value('--salesupdate--', 'Updating Files')
    mod_sku_list = ''
    for sku in reduced_sku_list[:-1]:
        mod_sku_list+=sku+' , '
    mod_sku_list+=sku_list[-1]
    options=webdriver.ChromeOptions()
    options.headless = True
    prefs={"download.default_directory":os.getcwd()+r'\exports'}
    options.add_experimental_option("prefs",prefs)
    driver = webdriver.Chrome(executable_path=chrome_path,chrome_options=options)
    driver.get("https://df.cwa.sellercloud.com/Inventory/PredictedPurchasing.aspx")
    SC_login(driver)
    sleep(2)
    driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$txtProductID").send_keys(mod_sku_list)
    driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ddlDaysOfOrder").send_keys(30)
    driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ddlDaysToOrder").send_keys(30)
    driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$btnSearch").click()
    driver.find_element(By.NAME, "ctl00$ContentPlaceHolder1$ddlAction").send_keys('Export to Excel')
    driver.find_element(By.NAME,"ctl00$ContentPlaceHolder1$btnDoAction").click()
    while True:
        try:
            shutil.move(path +rf'\exports\PredictPurchasing.xlsx' , path + r'\exports\Velocities.xlsx')
            break
        except: sleep(1)
    driver.close()
### totals sales of items in sku_list, per day, oer a period of time (dates list). Returns in DataFrame ###

def Adjust_SC_inv(driver,sku, qty,reason):
    driver.get(f"https://df.cwa.sellercloud.com/Inventory/ProductWareHouse.aspx?Id={sku}")
    
    before=   driver.find_element(By.ID, "ContentPlaceHolder1_ContentPlaceHolder1_grdSummary_ctl00__1")
    before=before.text.split(' ')[-1]
    driver.find_element(By.CLASS_NAME, "rpExpandHandle").click()
    driver.find_element(By.NAME, "ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolder1$pnlBarMain$i0$i0$Product_Warehouse_AdjustInventory1$txtQtyToAdjust").send_keys(qty)
    driver.find_element(By.NAME, "ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolder1$pnlBarMain$i0$i0$Product_Warehouse_AdjustInventory1$ddlWarehouseAdjustment").send_keys('542 Monterey Pass')
    driver.find_element(By.NAME, "ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolder1$pnlBarMain$i0$i0$Product_Warehouse_AdjustInventory1$ddlReason").send_keys('other')
    driver.find_element(By.NAME, "ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolder1$pnlBarMain$i0$i0$Product_Warehouse_AdjustInventory1$txtReason").send_keys(reason)
    driver.find_element(By.NAME, "ctl00$ctl00$ContentPlaceHolder1$ContentPlaceHolder1$pnlBarMain$i0$i0$Product_Warehouse_AdjustInventory1$btnAddAdjustment").click()
    sleep(2)
    after=   driver.find_element(By.ID, "ContentPlaceHolder1_ContentPlaceHolder1_grdSummary_ctl00__1")
    after=after.text.split()[-1]
    
    return before,after,qty
def wholesale_current(driver, sku):
    driver.get(f"https://df.cwa.sellercloud.com/Orders/AddItemsToOrder.aspx?OrderId=7808146")
    for n in range(4,136,2):
        a=driver.find_element(By.ID, f"grdItems_ctl00_ctl{'0'+ str(n) if n<10 else n}_hypProductID")
        if a.get_property('text')==sku:
            # driver.find_element(By.ID, f"grdItems_ctl00_ctl{'0'+ str(n) if n<10 else n}_txtQty").send_keys(Keys.BACKSPACE +Keys.BACKSPACE+ str(new_safety))
            return driver.find_element(By.ID, f"grdItems_ctl00_ctl{'0'+ str(n) if n<10 else n}_txtQty").get_property('value')
            break
    driver.find_element(By.ID , 'ddlAction').send_keys('Update Items Qty/Prices')
    driver.find_element(By.ID , 'btnDoAction').click()



def edit_wholesale(driver, sku, new_safety):
    if driver.current_url != "https://df.cwa.sellercloud.com/Orders/AddItemsToOrder.aspx?OrderId=7808146":
        driver.get(f"https://df.cwa.sellercloud.com/Orders/AddItemsToOrder.aspx?OrderId=7808146")
    for n in range(4,136,2):
        a=driver.find_element(By.ID, f"grdItems_ctl00_ctl{'0'+ str(n) if n<10 else n}_hypProductID")
        if a.get_property('text')==sku:
            old_safety = driver.find_element(By.ID, f"grdItems_ctl00_ctl{'0'+ str(n) if n<10 else n}_txtQty").get_property('value')
            driver.find_element(By.ID, f"grdItems_ctl00_ctl{'0'+ str(n) if n<10 else n}_txtQty").send_keys(Keys.BACKSPACE +Keys.BACKSPACE+ str(new_safety))
            return old_safety
def save_wholesale(window, driver):
    if driver.current_url != "https://df.cwa.sellercloud.com/Orders/AddItemsToOrder.aspx?OrderId=7808146":
        window.write_event_value('--save--',['Wholesale Order did not save\n      no changes have been made\n\n',False])
    
    else:
        driver.find_element(By.ID , 'ddlAction').send_keys('Update Items Qty/Prices')
        driver.find_element(By.ID , 'btnDoAction').click()
        window.write_event_value('--save--',['Wholesale Order Saved\n\n',True])