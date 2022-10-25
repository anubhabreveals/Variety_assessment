from pathlib import Path

import PySimpleGUI as sg
from openpyxl import load_workbook
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)



layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Action', size=(15,1)), sg.InputText(key='Action')],
    [sg.Text('State Change', size=(15,1)), sg.InputText(key='State Change')],
    [sg.Text('Phenomena', size=(15,1)), sg.InputText(key='Phenomena')],
    [sg.Text('Physical effect', size=(15,1)), sg.InputText(key='Physical effect')],
    [sg.Text('oRgan', size=(15,1)), sg.InputText(key='oRgan')],
    [sg.Text('Part', size=(15,1)), sg.InputText(key='Part')],
    [sg.Text('Input', size=(15,1)), sg.InputText(key='Input')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit(), sg.Button('Calculate')]
]

window = sg.Window('Idea Variety', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
    if event == 'Calculate':
        wb = load_workbook('data_Entry.xlsx')
        ws = wb.active
        column_action = ws['A']
        for cell in column_action:
            print(cell.value)

window.close()