
import threading
from Inventory_functions import *
import win32gui, win32con

### hides console window ###
# hide = win32gui.GetForegroundWindow()
# win32gui.ShowWindow(hide, win32con.SW_HIDE)

Queue=queue.Queue()
sg.theme('DarkTeal12')


def check_threads(thread_list,window):
    global running, disabled
    D_count=0
    for D in thread_list:
        D_count+=1
        if (D.is_alive()) and (disabled==False):
            disabled=True
            window['-TRANSFERS-'].update(disabled=disabled)
            window['-CONTAINERS-'].update(disabled=disabled)
            window['-SCDIRECT-'].update(disabled=disabled)
            running=D
            break
        elif (disabled==True) and (not D.is_alive()) and (D_count==len(thread_list)):
            disabled=False 
            window['-TRANSFERS-'].update(disabled=disabled)
            window['-CONTAINERS-'].update(disabled=disabled)
            window['-SCDIRECT-'].update(disabled=disabled)
            
            break
def BG_update_inv(window,download_queue):
    n=0
    D1.start()
    while D1.is_alive():
        n+=1
        if n==6:
            n=0
        sleep(.1)
        window.write_event_value('--invupdate--', ('Updating Files',n))
    D2.start()
    while D2.is_alive():
        n+=1
        if n==6:
            n=0
        sleep(.1)
        window.write_event_value('--invupdate--', ('Updating Files',n))
    n=0
    ### error message if failure ###
    
    error = download_queue.empty()
    if error:
        window.write_event_value('--invupdate--', ('Files Updated!\nUpdating General Inventory Sheets',n))
        low_inv_obj,sfty_obj= find_general_inv_files(r'exports/Inventory.xlsx', r'exports/safeties.xlsx')
        update_back_in_stock(low_inv_obj)
        update_inv_safety(low_inv_obj, sfty_obj)
        window.write_event_value('--invupdate--', ('Files Updated!\nGeneral Inventory Sheets Updated!',n))
    else:
        failure = download_queue.get()
        window.write_event_value('--invupdate--', (failure,0))

def loadup():
    window0=sg.Window('startup',[[sg.T(key='--LOADING--', justification='center', )]])
    startup=True
    n=0
    while True:
        event, values = window0.read(timeout=200)
        n+=1
        if n==3:
            n=0
        window0['--LOADING--'].update('loading'+('.'*n))
        if event == sg.WIN_CLOSED or event == 'Exit':
            window0.close()
            break
        # if event=='--LOADING--':
        #     window['--LOADING--'].update(value=values['--LOADING--'])
        if startup:
            startup=False
            init=threading.Thread(target=Initialize, args=(Queue, ))
            init.start()
        if not init.is_alive():
            window0.close()
            break

def main():

    global D1, D2, D3, D4, D5, D6, D7, D8, D9, disabled, dotcount


    general_layout = [[sg.Text('General Inventory')],
                [sg.Button('Update General Inventory', size = 190, key='-UPDATEINV-')],
                [sg.Text(key='-invupdate-', justification='center')]]

    transfers_layout = [[sg.Text('Warehouse Transfers')],
                [sg.Button('Update Containers to Warehouse Transfers', size = 190, key='-UPDATETRANS-')],
                [sg.Text('Containers: '),sg.Text(key='containers update list')]]

    containers_layout =[[sg.Button('back'),sg.Text('Containers Logs')],[sg.T('')],
                [sg.Text('Container Number'),sg.Text('                                                  '),sg.Text('Freight Forwarder'),sg.Text('                                                  '),sg.Text('ETA')],
                [sg.Input(key='Container Number'),sg.Input(key='Freight Forwarder'),sg.Input(key='ETA')],
                [sg.Text('Contents')],
                [sg.Input(key='Contents')],
                [sg.Text('MNFCR'),sg.Text('                                                                '),sg.Text('Management Notes'),sg.Text('                                               '),sg.Text('Additional Notes')],
                [sg.Input(key='MNFCR'),sg.Input(key='Management Notes'),sg.Input(key='Additional Notes')],
                [sg.T('')],
                [sg.Button('Add Container'), sg.Button('Search Containers')]]

    containers_search_layout =[[sg.Text('Search Containers')],
                [sg.Text('search sku: '), sg.Combo(reduced_sku_list,key='sku'),sg.Button('Search')],
                [sg.MLine(size=(80, 12),justification='left', k='-ML-', reroute_stdout=True,write_only=True, autoscroll=True, auto_refresh=True, key='QTY')]]

    ADJ_INV_layout = [[sg.Combo(reduced_sku_list,key='-ADJSKU-'),sg.Input('QTY',key='-ADJQTY-'),sg.Input('REASON',key='-ADJREASON-')],
                [sg.B('Update SC Inventory',key='-ADJINV-')],
                [sg.MLine(size=(80, 12),justification='left', reroute_stdout=True,write_only=True, autoscroll=True, auto_refresh=True, key='SCUPDATEDETAILS1')]]
    UPDATE_SFTY_layout = [[sg.Combo(reduced_sku_list,key='-SFTYSKU-',enable_events=True),sg.Input('QTY',key='-SFTYQTY-')],
                [sg.B('Update Safety Qtys',key='-UPDATESFTY-',expand_x=True),sg.B('Reset Safety', key='-SETSFTY-',expand_x=True), sg.B('Set Min', key='-SETMIN-',expand_x=True)],
                [sg.MLine(size=(80, 12),justification='left', reroute_stdout=True,write_only=True, autoscroll=True, auto_refresh=True, key='SCUPDATEDETAILS2')],
                [sg.B('save',key='-save-',expand_x=True)]]

    SC_direct_layout= [[sg.TabGroup([[sg.Tab('Adjust Inventory', ADJ_INV_layout), sg.Tab('Update Safety', UPDATE_SFTY_layout)]],expand_x=True, expand_y=True)]]
    

    tabgrp = [[sg.TabGroup([[sg.Tab('General', general_layout, key='-GENERAL-'),sg.Tab('Transfers', transfers_layout,element_justification='center', key='-TRANSFERS-'),sg.Tab('Containers', containers_search_layout, key='-CONTAINERS-', element_justification='center'),sg.Tab('SC Direct', SC_direct_layout, key='-SCDIRECT-', element_justification='center')]], expand_x=True, expand_y=True)]]

    window = sg.Window('Forecast Manager', tabgrp, finalize=True, resizable=True, element_justification="right", size=(700, 500))
    download_queue = queue.Queue()
    D1=Thread(target=download_inv, args=(window,download_queue))
    D2=Thread(target=download_safeties, args=(window,download_queue))
    BG_inv=Thread(target=BG_update_inv, args=(window,download_queue))
    D_list=[BG_inv]
    dotcount=0
    disabled=False
    scupdate_str=''
    Mline2=''
    saved=False


    while True:

        event, values = window.read(timeout=200, timeout_key='-TIMEOUT-')

        if event == '-TIMEOUT-':
            check_threads(D_list,window)
        if event == '--invupdate--':
            window['-invupdate-'].update(value=(values['--invupdate--'][0]+('.'*values['--invupdate--'][1])))
            window.refresh()
        
        
        if event == sg.WIN_CLOSED:
            sheet.values().update(spreadsheetId=IIMCHANGELOG_ID, range="'inventory adjustments'!A1", valueInputOption='USER_ENTERED', body={'values':change_log_i}).execute()
            if not saved:
                ok = sg.popup_ok('window closed without saving')
                if ok =='OK':
                    break
            else:    
                window.close()
                break

        if event == '-UPDATEINV-':
            if not BG_inv.is_alive():
                BG_inv.start()
                check_threads(D_list,window)

        if event ==  '-UPDATETRANS-':
            continue
            global containers,inv_safety, transfers
            containers,inv_safety, transfers = read_sheets()
            container_str = ''
            container_list = update_transfer() 
            for code in container_list:
                container_str+=code+',  '
            window['containers update list'].update(value=container_str) 

        if event == 'Search':
            index = search_containers(values['sku'].upper())
            cont_lst=f'{values["sku"].upper()}:  \n'
            for key in index.keys():
                cont_lst+=f'{key}: {index[key]} \n'
            window['QTY'].update(value=cont_lst)

        if event == '-ADJINV-':
            before,after,qty = Adjust_SC_inv(SC_driver, values['-ADJSKU-'], values['-ADJQTY-'],values['-ADJREASON-'])
            scupdate_str=f'{values["-ADJSKU-"]}:\n   value=Before adjustment: {before}\n   After Adjustment: {after}\n   Adjustment size: {qty}\n\n\n'
            window['SCUPDATEDETAILS1'].update(scupdate_str)
            change_log_i.append([values['-ADJSKU-'],qty,before,after,str(date.today()),signature])


        if event == '-SFTYSKU-':
            window['-SFTYQTY-'].update(value=wholesale_current(SC_driver,values['-SFTYSKU-']))
            window.refresh()
        if event == '-UPDATESFTY-':
            try:
                old=edit_wholesale(SC_driver, values["-SFTYSKU-"],values["-SFTYQTY-"])
                change_log_s.append([values["-SFTYSKU-"],old,values["-SFTYQTY-"],str(date.today()), signature])
                Mline2 += f'{values["-SFTYSKU-"]} Set Custom: \n{old} -> {values["-SFTYQTY-"]}\n\n'
                window['SCUPDATEDETAILS2'].update(value=Mline2)
                window.refresh()
            except:
                Mline2+=f'Please Choose a Valid SKU\n\n'
                window['SCUPDATEDETAILS2'].update(value=Mline2)
        if event=='-SETSFTY-':
            try:
                old=edit_wholesale(SC_driver, values["-SFTYSKU-"],sku_details[values["-SFTYSKU-"]]["safety"])
                change_log_s.append([values["-SFTYSKU-"],old,sku_details[values["-SFTYSKU-"]]["safety"],str(date.today()), signature])
                Mline2+=f'{values["-SFTYSKU-"]} Safety Reset: \n{old} -> {sku_details[values["-SFTYSKU-"]]["safety"]}\n\n'
                window['SCUPDATEDETAILS2'].update(value=Mline2)
                window.refresh()
            except:
                Mline2+=f'Please Choose a Valid SKU\n\n'
                window['SCUPDATEDETAILS2'].update(value=Mline2)
        if event=='-SETMIN-':
            try:
                old=edit_wholesale(SC_driver, values["-SFTYSKU-"],sku_details[values["-SFTYSKU-"]]["min"])
                change_log_s.append([values["-SFTYSKU-"],old,sku_details[values["-SFTYSKU-"]]["min"],str(date.today()), signature])
                Mline2 += f'{values["-SFTYSKU-"]} Set to Min: \n{old} -> {sku_details[values["-SFTYSKU-"]]["min"]}\n\n'
                window['SCUPDATEDETAILS2'].update(value=Mline2)
                window.refresh()
            except:
                Mline2+=f'Please Choose a Valid SKU\n\n'
                window['SCUPDATEDETAILS2'].update(value=Mline2)

        if event == "-save-":    
            sheet.values().update(spreadsheetId=IIMCHANGELOG_ID, range="'wholesale safeties changes'!A1", valueInputOption='USER_ENTERED', body={'values':change_log_s}).execute()
            save_wholesale(window,SC_driver)
            
        if event == "--save--":
            Mline2 += values['--save--'][0]
            window['SCUPDATEDETAILS2'].update(value=Mline2)
            saved=values['--save--'][1]

if __name__=='__main__':
    loadup()
    (SCOPES,SERVICE_ACCOUNT_FILE,service , CONTAINERS_ID ,TRANSFERS_ID,INV_SAFETY_ID,FORECAST_ID,IIMCHANGELOG_ID,path,sheet,sku_list,reduced_sku_list,sku_dict,sku_details,change_log_s,change_log_i,period_to_weeks,daterange,forecast_list,today,containers,inv_safety, transfers,SC_driver) = Queue.get()
    main()
    SC_driver.close()
























