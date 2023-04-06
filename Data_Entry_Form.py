# -*- coding: utf-8 -*-
"""
Created on Sun Mar 12 12:50:43 2023

@author: David Blackman
"""

#Import libraries
import PySimpleGUI as sg
import pandas as pd
import datetime


# Creating Timestamp
now = datetime.datetime.now()
t = now.strftime("%m/%d/%Y, %H:%M:%S")

EXCEL_FILE = 'Timekeeping.xlsx'
df = pd.read_excel(EXCEL_FILE)

# Add a theme
sg.theme('TealMono')

# Create Layout
layout = [
    [sg.Text('Please fill out the following fields.')],
    [sg.Text('ID', size = (4,1)), sg.InputText(key='ID')],
    [sg.Text('Product', size=(15,1)), sg.Combo(['Puzzles', 'Banks', 'Chalkboards', 'Tic Tac Toe','Post it Holder'], key='Product')],
    [sg.Text('Operation', size= (15,1)), sg.Combo(['Router', 'Laser', 'Print'], key = "Operation")],
    [sg.Checkbox('Start', size = (5,1), key = 'Start')], 
    [sg.Checkbox('End', size = (3,1), key = 'End')],
    [sg.Text('Date:', size = (4,1)), sg.Input(t, key = 'Date')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
    ]

#Create window
window = sg.Window('Timekeeping', layout)

#Clear Input 
def clear_input():
    for key in values:
        window[key]('')
        
    return None

#Finish entry
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