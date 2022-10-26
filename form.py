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
    [sg.Text('Concept ID', size=(15,1)), sg.Combo(['Concept 1', 'Concept 2', 'Concept 3', 'Concept 4', 'Concept 5', 'Concept 6', 'Concept 7', 'Concept 8', 'Concept 9', 'Concept 10'], key='Concept ID')],
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

def remove_duplicates(duplist):
    noduplist = []
    for element in duplist:
        if element not in noduplist:
            noduplist.append(element)

    return noduplist

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
        i=1
        j=2
        num_i = 'Concept '+str(i)
        num_j = 'Concept '+str(j)
        df_new_i = df[df['Concept ID'] == num_i]
        df_new_j = df[df['Concept ID'] == num_j]

        list_a = list(df_new_i['Action']) + list(df_new_j['Action'])
        u_a = len(remove_duplicates(list_a))
        n_a = len(list_a)

        list_s = list(df_new_i['State Change']) + list(df_new_j['State Change'])
        u_s = len(remove_duplicates(list_s))
        n_s = len(list_s)

        list_ph = list(df_new_i['Phenomena']) + list(df_new_j['Phenomena'])
        u_ph = len(remove_duplicates(list_ph))
        n_ph = len(list_ph)

        list_e = list(df_new_i['Physical effect']) + list(df_new_j['Physical effect'])
        u_e = len(remove_duplicates(list_e))
        n_e = len(list_e)

        list_r = list(df_new_i['oRgan']) + list(df_new_j['oRgan'])
        u_r = len(remove_duplicates(list_r))
        n_r = len(list_r)

        list_p = list(df_new_i['Part']) + list(df_new_j['Part'])
        u_p = len(remove_duplicates(list_p))
        n_p = len(list_p)

        list_i = list(df_new_i['Input']) + list(df_new_j['Input'])
        u_i = len(remove_duplicates(list_i))
        n_i = len(list_i)


        print(u_a,n_a)
        print(df_new_i)
        print(df_new_j)

window.close()