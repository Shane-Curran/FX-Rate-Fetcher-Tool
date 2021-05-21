import yfinance as yf
import tkinter
import PySimpleGUI as sg
import subprocess, os, platform



if __name__ == '__main__':
    #===========================================================================================================
    # GUI Formatting
    #===========================================================================================================
    items = ['EUR/USD','USD/JPY','GBP/USD','AUD/USD','NZD/USD','USD/CNY','USD/HKD','USD/SGD','USD/INR','USD/MXN',
                'USD/PHP','USD/IDR','USD/THB','USD/MYR','USD/ZAR','USD/RUB']
    cur_dictionary = {"EUR/USD" : "EURUSD=X","USD/JPY" : "JPY=X","GBP/USD" : "GBPUSD=X","AUD/USD" : "AUDUSD=X",
                        "NZD/USD" : "NZDUSD=X","USD/CNY" : "CNY=X","USD/HKD" : "HKD=X","USD/SGD" : "SGD=X",
                        "USD/INR" : "INR=X","USD/MXN" : "MXN=X","USD/PHP" : "PHP=X","USD/IDR" : "IDR=X",
                        "USD/THB" : "THB=X","USD/MYR" : "MYR=X","USD/ZAR" : "ZAR=X","USD/RUB" : "RUB=X"}
    selections = []
    path = ""
    start_date = ""
    end_date = ""
    cur_list = ""

    sg.theme("DarkTeal2")
    layout = [[sg.Text('Select currencies below:')],
                [sg.Listbox(values=items, size=(40,20), enable_events=True, bind_return_key=True, select_mode='single', key='_LISTBOX_')],
                [sg.Checkbox('Select Multiple',default=False,enable_events=True,key='_SELECT_MULTI_')],
                [sg.Text('Choose start date:'), sg.Input(key='-IN3-', size=(20,1)), sg.CalendarButton('Start Date', close_when_date_chosen=True,  target='-IN3-', location=(0,0), no_titlebar=False)],
                [sg.Text('Choose end date:'), sg.Input(key='-IN4-', size=(20,1)), sg.CalendarButton('End Date', close_when_date_chosen=True,  target='-IN4-', location=(0,0), no_titlebar=False)],
                [sg.T("")], [sg.Text("Choose a folder: "), sg.Input(key="-IN2-" ,change_submits=True), sg.FolderBrowse(key="-IN-")],
                [sg.Button("Submit"), sg.Button('Quit')]]
    window = sg.Window('FX Rates by Shane Curran', layout,size= (1000,600), location = (0,0))
    while True:
        event, values = window.Read()
        if event is None:
            exit()
        elif event == '_LISTBOX_':
            #print(values['_LISTBOX_'])
            selections = values['_LISTBOX_']
        elif event == '_SELECT_MULTI_':
            if values['_SELECT_MULTI_'] is True:
                window.Element('_LISTBOX_').Update(select_mode='multiple') # breaks, of course
                #print('now you can select multiple things!')
            else:
                window.Element('_LISTBOX_').Update(select_mode='single') # breaks, of course
                #print('now you can select one at a time')
        elif event == "Submit":
            #print(values["-IN-"])
            path = values["-IN-"]
            start_date = values['-IN3-']
            start_date = start_date[0:10]
            end_date = values['-IN4-']
            end_date = end_date[0:10]
            break
        elif event == 'Quit':
            exit()

    window.Close()

    print(selections)
    print(start_date)
    print(end_date)
    print(path)

    for i in selections:
        cur_list = cur_list + " " + cur_dictionary[i]

    if cur_list[0] == " ":
        cur_list = cur_list[1:]

    print(cur_list)

data = yf.download(str(cur_list), start= start_date, end= end_date, group_by= 'ticker')

path = path + '/FX-Output.xlsx'
print(path)

data.to_excel(path)


if platform.system() == 'Darwin':       # macOS
    subprocess.call(('open', path))
elif platform.system() == 'Windows':    # Windows
    os.startfile(path)
else:                                   # linux variants
    subprocess.call(('xdg-open', path))