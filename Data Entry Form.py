# -*- coding: utf-8 -*-
"""
Created on Sun Mar 12 12:50:43 2023

@author: David Blackman
"""


import PySimpleGUI as sg
import pandas as pd


EXCEL_FILE = 'Timekeeping.xlsx'
df = pd.read_excel(EXCEL_FILE)

# Add a theme
sg.theme('DarkTeal9')

# Create Layout
layout = [
    [sg.Text('Please fill out the following fields.')],
    [sg.Text('ID', size = (4,1)), sg.InputText(key='ID')],
    [sg.Text('Product', size=(15,1)), sg.Combo(['Puzzles', 'Banks', 'Chalkboards'], key='Product')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
    ]

#Create window
window = sg.Window('Timekeeping', layout)

def clear_input():
    for key in values:
        window[key]('')
        
    return None
#
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    
    if event == 'Clear':
        clear_input()
    
    if event == 'Submit':
        df = df.append(values, ignore_index = True)
        df.to_excel(EXCEL_FILE, index = False)
        sg.popup('Data Saved!')
        clear_input()
        
window.close()