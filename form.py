from pathlib import Path
from tabulate import tabulate

import PySimpleGUI as sg
from openpyxl import load_workbook
import pandas as pd
import numpy as np

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
    [sg.Submit(), sg.Button('Clear'), sg.Exit(), sg.Button('Calculate Variety'), sg.Button('Show Database'), sg.Button('Clear Database')]
]

# layout2 = [[sg.Multiline('', size=(80,10), key='database')]]

window = sg.Window('Idea Variety', layout)

def database_window():
    headings = ['Concept ID','Action','State Change','Phenomena','Physical effect','oRgan','Part','Input']
    df1=df.values.tolist()
    database_layout = [[sg.Table(values = df1, headings = headings,
    auto_size_columns=True, justification='left')]]

    database_window = sg.Window('Database', database_layout, modal=True)
    while True:
        event, values = database_window.read()
        if event == sg.WIN_CLOSED:
            break
    database_window.close()



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
        if values['Concept ID']=='' or values['Action']=='' or values['State Change']=='' or values['Phenomena']=='' or values['Physical effect']=='' or values['oRgan']=='' or values['Part']=='' or values['Input']=='':
            sg.popup('All fields are required!')
        else:
            new_record = pd.DataFrame(values, index=[0])
            df = pd.concat([df, new_record], ignore_index=True)
            df.to_excel(EXCEL_FILE, columns=['Concept ID','Action','State Change','Phenomena','Physical effect','oRgan','Part','Input'], index=False)
            sg.popup('Data saved!')
            clear_input()
    if event == 'Calculate Variety':
        concept_list = list(df['Concept ID'])
        total_num_of_concepts = len(remove_duplicates(concept_list))
        n = total_num_of_concepts + 1
        d_ij_list = []
        for i in range(1,n):
            for j in range(1,n):
                if i != j:
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

                    d_ij = ((u_a/n_a)+(u_s/n_s)+(u_ph/n_ph)+(u_e/n_e)+(u_r/n_r)+(u_p/n_p)+(u_i/n_i))/7

                   # print(i,j)
                   # print(u_a,n_a)
                   # print(d_ij)
                   # print(df_new_i)
                   # print(df_new_j)
                    d_ij_list.append(round(d_ij, 4))
                else:
                    d_ij = 0
                    d_ij_list.append(round(d_ij, 4))
     
        d_ij_matrix = []
        while d_ij_list != []:
            d_ij_matrix.append(d_ij_list[:n-1])
            d_ij_list = d_ij_list[n-1:]
        
        d_ij_matrix = np.array(d_ij_matrix)
        d_ij_sum = d_ij_matrix.sum(axis = 1)  
        for line in d_ij_matrix:
            print ('  '.join(map(str, line)))
        v_i_list = []
        for i in range(1,n):
            v_i = (d_ij_sum[i-1])/(n-2)
            v_i_list.append(v_i)
            text_print_1 = 'V['+str(i)+'] = '+str(round(v_i, 4)) 
            print(text_print_1)
        v_cs = (sum(v_i_list))/(n-1)
        sg.popup('Variety score of the solution space is: '+str(round(v_cs, 4)), title='Variety Score')
        
    if event == 'Show Database':
        print(tabulate(df, headers = 'keys', tablefmt = 'psql', showindex=False))
        #sg.PopupScrolled(df, size=(80,10))
        database_window()
    if event == 'Clear Database':
        df.drop(df.index, inplace=True)
        df.to_excel(EXCEL_FILE, index=False)
        print(tabulate(df, headers = 'keys', tablefmt = 'psql', showindex=False))
window.close()