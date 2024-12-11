import PySimpleGUI as sg
import pandas as pd

sg.theme('DarkTeal9')

EXCEL_FILE = 'data_entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields: ')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],    
    [sg.Text('Favourite Color', size=(15,1)), sg.Combo(['Green', 'Red', 'Blue'], key='Favourite Color')],
    [sg.Text('I Speak', size=(15,1)),
        sg.Checkbox('Urdu', key='Urdu'),
        sg.Checkbox('English', key='English'),
        sg.Checkbox('Punjabi', key='Punjabi')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0, 10)], initial_value=0, key='Children')],

    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple Data Entry Form', layout)

def clear_inputs():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_inputs()
    if event == 'Submit':
        # print(event, values)
        new_df = pd.DataFrame([values])
        df = pd.concat([df, new_df], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Saved!')
        clear_inputs()

window.close() 