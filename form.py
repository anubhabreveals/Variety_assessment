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


def calculate_window(x):
    print(x)
    df = pd.read_excel(EXCEL_FILE)
    con_id_list = []
    for i in range(1,x+1):
        con_id_list.append('Concept '+str(i))
    calculate_layout = [
        [sg.Text('Please fill out the following fields:')],
        [sg.Text('Concept ID', size=(15,1)), sg.Combo(con_id_list, key='Concept ID')],
        [sg.Text('Action', size=(15,1)), sg.InputText(key='Action')],
        [sg.Text('State Change', size=(15,1)), sg.InputText(key='State Change')],
        [sg.Text('Phenomena', size=(15,1)), sg.InputText(key='Phenomena')],
        [sg.Text('Physical effect', size=(15,1)), sg.InputText(key='Physical effect')],
        [sg.Text('oRgan', size=(15,1)), sg.InputText(key='oRgan')],
        [sg.Text('Part', size=(15,1)), sg.InputText(key='Part')],
        [sg.Text('Input', size=(15,1)), sg.InputText(key='Input')],
        [sg.Submit(), sg.Button('Clear'), sg.Exit(), sg.Button('Calculate Variety'), sg.Button('Show Database'), sg.Button('Clear Database')]
    ]

    calculate_window = sg.Window('VariAnT v1.0', calculate_layout)

    def clear_input():
        for key in values:
            calculate_window[key]('')
        return None
    def warning_window():
        warning_layout = [[sg.Text('Do you really want to delete these records? \nThis process cannot be undone')],[sg.Button('Cancel'), sg.Button('Delete', button_color='red')]]

        warning_window = sg.Window('', warning_layout, modal=True)
        while True:
            event, values = warning_window.read()
            if event == sg.WIN_CLOSED or event == 'Cancel':
                break
            if event == 'Delete':
                df.drop(df.index, inplace=True)
                df.to_excel(EXCEL_FILE, index=False)
                print(tabulate(df, headers = 'keys', tablefmt = 'psql', showindex=False))
                sg.Popup('Done!')
                break
                
        warning_window.close()

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

    def remove_duplicates(duplist):
        noduplist = []
        for element in duplist:
            if element not in noduplist:
                noduplist.append(element)

        return noduplist


    def n_k_minus_s_k(list1,list2):
        union = set(list1).union(set(list2))
        intersection = set(list1).intersection(set(list2))
        n_k_min_s_k = len(list1)+len(list2)-2*len(intersection)
        return n_k_min_s_k


    while True:
        event, values = calculate_window.read()
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

                        list_a_i = list(df_new_i['Action'])
                        list_a_j = list(df_new_j['Action'])
                        list_a = list_a_i + list_a_j
                        u_a = n_k_minus_s_k(list_a_i,list_a_j)
                        n_a = len(list_a)

                        list_s_i = list(df_new_i['State Change'])
                        list_s_j = list(df_new_j['State Change'])
                        list_s = list_s_i + list_s_j
                        u_s = n_k_minus_s_k(list_s_i,list_s_j)
                        n_s = len(list_s)

                        list_ph_i = list(df_new_i['Phenomena'])
                        list_ph_j = list(df_new_j['Phenomena'])
                        list_ph = list_ph_i + list_ph_j
                        u_ph = n_k_minus_s_k(list_ph_i,list_ph_j)
                        n_ph = len(list_ph)

                        list_e_i = list(df_new_i['Physical effect'])
                        list_e_j = list(df_new_j['Physical effect'])
                        list_e = list_e_i + list_e_j
                        u_e = n_k_minus_s_k(list_e_i,list_e_j)
                        n_e = len(list_e)

                        list_r_i = list(df_new_i['oRgan'])
                        list_r_j = list(df_new_j['oRgan'])
                        list_r = list_r_i + list_r_j
                        u_r = n_k_minus_s_k(list_r_i,list_r_j)
                        n_r = len(list_r)

                        list_p_i = list(df_new_i['Part'])
                        list_p_j = list(df_new_j['Part'])
                        list_p = list_p_i + list_p_j
                        u_p = n_k_minus_s_k(list_p_i,list_p_j)
                        n_p = len(list_p)

                        list_i_i = list(df_new_i['Input'])
                        list_i_j = list(df_new_j['Input'])
                        list_i = list_i_i + list_i_j
                        u_i = n_k_minus_s_k(list_i_i,list_i_j)
                        n_i = len(list_i)

                        d_ij = ((u_a/n_a)+(u_s/n_s)+(u_ph/n_ph)+(u_e/n_e)+(u_r/n_r)+(u_p/n_p)+(u_i/n_i))/7

                        #print(i,j)
                        #print(u_s,n_s)
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
            warning_window()
    calculate_window.close()

layout = [[sg.Text('Total no. of concepts:', size=(15,1)), sg.Spin([i for i in range(1,100)], size=(5,1), initial_value=1, key='con_num')],[sg.Button('Start'),sg.Button('Exit')]]
window = sg.Window('VariAnT v1.0', layout)
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Start':
        con_count = values['con_num']
        calculate_window(con_count)
    
window.close()