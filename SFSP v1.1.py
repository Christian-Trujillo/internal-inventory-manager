import matplotlib.pyplot as plt
from ForecastFunctions import *
import time
import sys
import threading
import queue as Q
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import sklearn
import win32gui, win32con

### do not use tkinter, does not support threads ###
# chars = r"/—\|" #

### hides console window ###
# hide = win32gui.GetForegroundWindow()
# win32gui.ShowWindow(hide, win32con.SW_HIDE)


THREAD_EVENT = '-THREAD-'
sg.theme('DarkBlue')


def INIT():
    ''' display loading screen while start() function runs
        choose "import" to reup new data\n
        choose "skip" to import data from Forecast Google Sheet (faster)\n
        '''
    layout = [[sg.Text(key='startup_txt0')], [sg.Text(key='startup_txt1')], [sg.Text(key='startup_txt2')], [sg.Text(key='startup_txt3')], [
        sg.Text(key='startup_txt4')], [sg.Text(key='startup_txt5')], [sg.B('Import Data', key='-IMPORT-'), sg.B('skip', key='-SKIP-',)]]
    window = sg.Window('Forecast Manager', layout, finalize=True,
                       size=(500, 250), element_justification='c')
                       
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            window.close()
            break
        if event == THREAD_EVENT:
            window[f'startup_txt{values[THREAD_EVENT][2]}'].update(value=(
                values[THREAD_EVENT][0]+('.'*values[THREAD_EVENT][1])+(" "*(3-values[THREAD_EVENT][1]))))
            window.refresh()
        if event == '-IMPORT-':
            window['-SKIP-'].update(disabled=True)
            window['-IMPORT-'].update(disabled=True)
            start_thread = Thread(target=start, args=(window,))
            start_thread.start()
        if event == '-SKIP-':
            global reduced_sku_list, sku_list, period_to_weeks, sku_details, forecasts, unprocessed_sales, processed_sales, num_period, period_num, containers, transfers, inv_safety
            window.close()
            queue = Q.Queue()
            Initialize(queue)
            (reduced_sku_list, sku_list, period_to_weeks,
             sku_details, num_period, period_num, credentials) = queue.get()
            read_sheets(queue)
            (containers, transfers, inv_safety) = queue.get()
            forecasts = read_forecasts()
            processed_sales = read_sales()
            unprocessed_sales = processed_sales*0


def start(window):
    ''' updates all relevent info while updating loading screen, then runs main() once all threads are finished'''
    global reduced_sku_list, sku_list, period_to_weeks, sku_details, containers, transfers, inv_safety, unprocessed_sales, processed_sales, forecasts, num_period, period_num
    queue = Q.Queue()
    loading_text = [' Initializing', ' Importing Sheets',
                    ' Importing Product Info', ' Processing Sales', ' Processing Forecasts']
    loaded_text = [' Initialized   ', ' Sheets imported   ', ' Product Info Imported   ',
                   ' Sales Processed   ', ' Forecasts Processed   ',' FINISHED!  ']
    n = 0
    for func in [Initialize, read_sheets, download_files, process_sales, Process_Forecast]:
        if func == Process_Forecast:
            done = Thread(target=func, args=(processed_sales, queue))
        else:
            done = Thread(target=func, args=(queue,))
        done.start()
        i = 0
        ### update loading screen each .1 second that thread is not finished ###
        while done.is_alive():
            time.sleep(.1)
            if i == 4:
                i = 0
            window.write_event_value('-THREAD-', (loading_text[n], i, n))
            i += 1
        ### defines variables returned by starting functions ###
        if not done.is_alive():
            if func == Initialize:
                (reduced_sku_list, sku_list,
                 period_to_weeks, sku_details, num_period, period_num, credentials) = queue.get()
            if func == read_sheets:
                (containers, transfers, inv_safety) = queue.get()
            if func == process_sales:
                (unprocessed_sales, processed_sales) = queue.get()
            if func == Process_Forecast:
                forecasts = queue.get()
            window[f'startup_txt{n}'].update(value=loaded_text[n])
            n += 1
    ### once all threads are finished, update screen with "finished" then run main() ###
    Export_Forecast(prepare_exports(processed_sales, forecasts))
    Days_in_stock()
    # forecasts=forecasts[:3]
    window['startup_txt5'].update(value=' Finished!  ')
    window.refresh()
    window.write_event_value('Exit', 'Exit')


def main():
    '''opens main GUI with:\n
        Sku Details Tab\n
        Forecast Table Tab\n
        and production Schedule Export Tab'''
    ### midpoint between AD and ML forecast ###
    mid_forecast = forecasts[0].add(forecasts[2])/2
    table_forecast = mid_forecast.reset_index().values.tolist()
    # table_forecast.insert(0,forecast.reset_index().values.tolist())
    column_names = mid_forecast.columns.tolist()
    column_names.insert(0, 'SKU')
    RCM = ['', ['Copy']]
    last_sku = None
    last_period = None
    table_values = [[]]
    table_headings = (processed_sales.index.tolist()[-5:])
    table_headings.insert(0, '                ')
    sg.set_options(font=("Courier New", 10))
    layout = [[sg.Combo(reduced_sku_list, key='chosen_sku'), sg.Combo(['1 Week', '2 Weeks', '1 Month', '2 Months', '3 Months', '4 Months', '5 Months', '6 Months', '7 Months', '8 Months', '9 Months'], key='chosen_period')],
              [sg.Text(key='sku_info', justification='left', background_color='lightgrey', text_color='black', expand_x=True, expand_y=True), sg.Canvas(key='figCanvas', size=(800, 800), expand_x=True, expand_y=True)],
              [sg.Table(values=table_values, headings=table_headings, auto_size_columns=False, col_widths=list(map(lambda x:len(x)+5, table_headings)), row_height=30, key='info table', justification='bottom', expand_y=True, expand_x=True, size=(10, 4))]]

    forecast_table = [[sg.Table(values=table_forecast,key='forecast table', headings=column_names, expand_x=True, expand_y=True, right_click_menu=RCM, select_mode=sg.TABLE_SELECT_MODE_EXTENDED)]]

    export_tab = [[sg.CBox('1', default=True, key='cbox1'), sg.CBox('2', default=True, key='cbox2'), sg.CBox('3', default=True, key='cbox3'), sg.CBox('4', default=True, key='cbox4'), sg.CBox('5', default=True, key='cbox5'), sg.CBox('6', default=True, key='cbox6'), sg.CBox('7', default=True, key='cbox7'), sg.CBox('8', default=True, key='cbox8'), sg.CBox('9', default=True, key='cbox9'), sg.CBox('10', default=True, key='cbox10'), sg.CBox('11', default=True, key='cbox11'), sg.CBox('"disregard"', key='cbox12')],
        [sg.T(' LEAD TIME'), sg.T('    '), sg.T('IDEAL QTY'),
         sg.T('    '), sg.T('MAX CONTAINERS')],
        [sg.Combo(list(period_num.keys()), key='--LEAD--'), sg.T('   '), sg.Combo(list(period_num.keys()),key='--IDEAL--'), sg.T('      '), sg.Combo(list(range(25)), key='--MAXCONTAINERS--')],
        [sg.T('')],
        [sg.FileBrowse(key='--PRODUCTION--'), sg.T('', key='--PTXT--')],
        [sg.T('', key="--EXPORTED?--", expand_x=True)],
        [sg.Button('Export Production Schedule', key='--EXPORT--', expand_x=True)]]

    tabgrp = [[sg.TabGroup([[sg.Tab('analyze by sku', layout), sg.Tab('forecast table', forecast_table), sg.Tab('export production schedule', export_tab)]], expand_x=True, expand_y=True)]]

    window = sg.Window('Forecast Manager', tabgrp, finalize=True, resizable=True, element_justification="right", size=(500, 500))
    Ftable:sg.Table = window['forecast table']
    window.bind("<Control-C>", "Copy")
    window.bind("<Control-c>", "Copy")
    while True:
        event, values = window.read(timeout=200)
        if values == None:
            break
        if values['--PRODUCTION--'] != '':
            window['--PTXT--'].update(value=values['--PRODUCTION--'])
            window.refresh()
        if event == sg.WIN_CLOSED or event == 'Exit':
            window.close()
            break
        elif (values['chosen_sku'] != '') & (values['chosen_period'] != '') & ((values['chosen_sku'] != last_sku) | (values['chosen_period'] != last_period)):
            '''if either of the 2 parameters change, update graph, sku_info and table'''
            try:
                sales_Df, errors = new_graph_forecasts(values['chosen_sku'], values['chosen_period'], forecasts[0], processed_sales,
                                                       sku_details[values['chosen_sku']]['forecastability'][1])[-(52+period_to_weeks[values['chosen_period']]):]
                xData = sales_Df.index.values.tolist()
                for n in range(len(xData)):
                    xData[n] = xData[n][:-5]
                yData = sales_Df.values
                # delete old graph
                try:
                    fig_agg.get_tk_widget().forget()
                    plt.close('all')
                except:
                    pass
                # Make and show plot
                fig = plt.figure(figsize=(10, 3.5),)
                plt.plot(xData, yData, color='black')
                ### plotting mltiple times plots many times??? ###
                # plt.plot(xData, [sales_Df['sales'].mean() for y in yData], color='gold')
                # plt.plot(errors.index.tolist(), [errors[0].mean for y in range(errors.shape[0])], color='dark gold' )
                plt.xticks(rotation=80)
                plt.fill_between(
                    xData[-errors.shape[0]:], errors['high error'], errors[0], color='lightblue')

                plt.fill_between(
                    xData[-errors.shape[0]:], errors['low error'], errors[0], color='darksalmon')
                fig_agg = draw_figure(window['figCanvas'].TKCanvas, fig)
                last_sku = values['chosen_sku']
                last_period = values['chosen_period']

                ### update sku info, then table ###
                update_sku_info(
                    window, mid_forecast, values['chosen_sku'], sku_details[values['chosen_sku']], processed_sales)
                table_values = [unprocessed_sales[last_sku].transpose().values.tolist()[-5:], processed_sales[last_sku].transpose().astype(int).values.tolist()[-5:], processed_sales[last_sku].transpose(
                ).rolling(4).mean().values.tolist()[-5:], (unprocessed_sales[last_sku]-processed_sales[last_sku].mean()).apply(lambda x:round(x, 0)).transpose().astype(int).values.tolist()[-5:]]
                index = ['Sales', 'Sales (fixed)',
                         'Running Average', 'Deviation']
                for n in range(4):
                    table_values[n].insert(0, index[n])
                window['info table'].update(values=table_values)

                window.refresh()
            except:
                pass
        elif event == 'Copy':
            try:
                table_selection=[(str(x) for x in Ftable.Values[row]) for row in Ftable.SelectedRows]
                table_selection.insert(0,Ftable.ColumnHeadings)
                text='\n'.join('\t'.join(cell for cell in row) for row in table_selection)
                window.TKroot.clipboard_clear()
                window.TKroot.clipboard_append(text)
            except Exception as e: print(e)
        
        
        ### export, then format with openpyxl ###
        ### make list of chosen factories to pass into PS function ###
        elif event == '--EXPORT--':
            chosen_factories = []
            for box in range(1, 13):
                if (box == 12) & (values[f'cbox{box}']):
                    chosen_factories.append('disregard')
                elif values[f'cbox{box}']:
                    chosen_factories.append(box)
            try:
                if (period_num[values['--LEAD--']]+period_num[values['--IDEAL--']] > 9) or (period_num[values['--LEAD--']]+period_num[values['--IDEAL--']] ==0):
                    raise ValueError('Please Enter a Valid Lead Time Total')
                period = num_period[period_num[values['--LEAD--']
                                               ]+period_num[values['--IDEAL--']]]

            except:
                period = '5 Months'
                window['--LEAD--'].update(value='1 Month')
                window['--IDEAL--'].update(value='4 Months')
            try:
                Export_Production_Schedule(
                    Prepare_Production_Schedule(sku_details, period, chosen_factories, values['--PRODUCTION--'], values['--MAXCONTAINERS--']))
                window['--EXPORTED?--'].update(value='-- FINISHED --')
                window.refresh()
            except  PermissionError:
                window['--EXPORTED?--'].update(
                    value='Permission Error: \nplease close export file and retry')
                window.refresh()
            except  KeyError:
                window['--EXPORTED?--'].update(
                    value='Incorrect File: \nplease choose a proper "In Production" workbook and retry')
                window.refresh()
            except FileNotFoundError:
                window['--EXPORTED?--'].update(
                    value='File Not Found: \nplease choose a proper "In Production" workbook and retry')
                window.refresh()


def draw_figure(canvas, figure):
    '''draws graph in SKU Details Tab'''
    figure_canvas_agg = FigureCanvasTkAgg(figure, canvas)
    figure_canvas_agg.get_tk_widget().pack(side='top', fill='both', expand=1)
    figure_canvas_agg.draw()
    return figure_canvas_agg


def update_sku_info(window, forecast, sku, sku_details, sales):
    '''updates details table in SKU Details Tab'''
    # desc = sku_details['desc']
    # sku_type ='skus item type here'
    size = sku_details['size']
    velocity = round(sales[sku].tail(4).sum()/30, 2)
    forecastability = sku_details['forecastability'][0] + \
        '; '+str(sku_details['forecastability'][1])
    week = forecast.loc[sku, '1 Week'].astype(int)
    two_week = forecast.loc[sku, '2 Weeks'].astype(int)
    month = forecast.loc[sku, '1 Month'].astype(int)
    two_month = forecast.loc[sku, '2 Months'].astype(int)
    three_month = forecast.loc[sku, '3 Months'].astype(int)
    four_month = forecast.loc[sku, '4 Months'].astype(int)
    five_month = forecast.loc[sku, '5 Months'].astype(int)
    six_month = forecast.loc[sku, '6 Months'].astype(int)
    seven_month = forecast.loc[sku, '7 Months'].astype(int)
    eight_month = forecast.loc[sku, '8 Months'].astype(int)
    nine_month = forecast.loc[sku, '9 Months'].astype(int)
    update_val = f'SKU: {sku}\n\nSize: H:{size[0]} L:{size[1]} W:{size[2]}\n\nVelocity: {velocity}\n\nForecastability: {forecastability}\n\n1 Week: {week} ± {int(week*sku_details["forecastability"][1]/100)}\n\n2 Weeks: {two_week} ± {int(two_week*sku_details["forecastability"][1]/100)}\n\n1 Month: {month} ± {int(month*sku_details["forecastability"][1]/100)}\n\n2 Months: {two_month} ± {int(two_month*sku_details["forecastability"][1]/100)}\n\n3 Months: {three_month} ± {int(three_month*sku_details["forecastability"][1]/100)}\n\n4 Months: {four_month} ± {int(four_month*sku_details["forecastability"][1]/100)}\n\n5 Months: {five_month} ± {int(five_month*sku_details["forecastability"][1]/100)}\n\n6 Months: {six_month} ± {int(six_month*sku_details["forecastability"][1]/100)}\n\n7 Months: {seven_month} ± {int(seven_month*sku_details["forecastability"][1]/100)}\n\n8 Months: {eight_month} ± {int(eight_month*sku_details["forecastability"][1]/100)}\n\n9 Months: {nine_month} ± {int(nine_month*sku_details["forecastability"][1]/100)}'
    window['sku_info'].update(value=update_val)


def new_graph_forecasts(sku, period, AD_forecast, sales, error):
    """create line for AD, create prediction from ML pickles, average and graph with new index"""
    ### create line for AD forecast ###
    P = period_to_weeks[period]
    N = AD_forecast.loc[sku, period]
    n1 = sales[sku][-1:][0]
    n1_date = sales.index[-1].split('/')
    n1_date = date(int(n1_date[2]), int(n1_date[0]), int(n1_date[1]))
    dates_list = []
    values_list = []
    ### create index for AD line with date formatting, and values based on last weeks sales and total sales over period ###
    for T in range(P+1):
        dates_list.append(
            (n1_date+relativedelta(weeks=T)).strftime("%m/%d/%Y"))
        values_list.append(n1+((2*T/P)*((N/P)-n1)))
    ### produce ML forecast ###
    S = pd.DataFrame(sales[f'{sku}'].values, columns=[
                     'Sales'], index=sales.index)
    S['Date'] = S.index
    for index in S.index:
        S.loc[index, 'Month'] = datetime.strptime(
            S.loc[index, 'Date'], '%m/%d/%Y').strftime('%m')
        S.loc[index, 'Year'] = datetime.strptime(
            S.loc[index, 'Date'], '%m/%d/%Y').strftime('%y')
    S = S.drop(['Date'], 1)
    S['Rolling Average'] = S['Sales'].rolling(3).mean()
    S['Rolling delta'] = S['Sales'].rolling(
        2).apply(lambda x: x.iloc[1] - x.iloc[0])
    S['avg delta'] = S['Rolling delta'].rolling(3).mean()
    S['Rolling delta 2'] = S['Rolling delta'].rolling(
        2).apply(lambda x: x.iloc[1] - x.iloc[0])
    S['avg delta 2'] = S['Rolling delta 2'].rolling(3).mean()
    S.fillna(0, inplace=True)
    S = S.astype(int)
    pickle_in = open(f'pickles\LR {sku}-{period}.pickle', 'rb')
    clf = pickle.load(pickle_in)
    mlvalues = list(clf.predict(np.array(S.tail(period_to_weeks[period]))))
    mlvalues.insert(0, values_list[0])
    ### average 2 forecasts +- error bar ###
    high_error = []
    low_error = []
    for val in range(len(values_list)):
        values_list[val] = (values_list[val]+mlvalues[val])/2
        if val == 0:
            high_error.append(values_list[val])
            low_error.append(values_list[val])
        else:
            high_error.append(values_list[val]+(values_list[val]*error/100))
            low_error.append(values_list[val]-(values_list[val]*error/100))

    df = pd.DataFrame(values_list, index=dates_list)
    df['high error'] = pd.DataFrame(high_error, index=dates_list)
    df['low error'] = pd.DataFrame(low_error, index=dates_list)

    sales.sku = sales[sku]
    for j in range(1, len(dates_list)):
        sales.sku = sales.sku.append(pd.Series(np.nan, index=[dates_list[j]]))

    sales.sku = pd.DataFrame(
        sales.sku.values, index=sales.sku.index.values.tolist(), columns=[sku])
    sales.sku = pd.merge(sales.sku, df, how='left',
                         left_index=True, right_index=True)
    sales.sku.columns = ["sales", "prediction", 'high error', 'low error']
    return sales.sku, df


if __name__ == '__main__':
    INIT()
    main()
