import PySimpleGUI as pg
import os
import pandas as pd
from pathlib import Path
import re


# Set theme
pg.theme("SystemDefault1")


# Create layout
file_list_column = [      
    [
        pg.Input(
        "",
        size=(33, 1), 
        pad=(5, (0, 10)),
        enable_events=True, 
        key="-INPUT-"
        ),
        pg.Button
        (
            "Search", 
            pad=(10, (0, 10)),
            size=(8, 1),
        )
    ], 
    [        
        pg.Listbox(            
            values=[],            
            enable_events=True,            
            size=(45, 25), 
            key="-FILE_LIST-",
        )    
    ]
]

file_viewer_column = [    
    [
        pg.Text
        (
            "Type Work Order number and Search", 
            size=(50, 2),
            pad=(0, (5, 10)),
            key="-DESC-"
        )
    ],    
    [
        pg.Text
        (
            "File path: ", 
            size=(50, 2),
            pad=(0, 0),
            key="-TOUT-"
        )
    ],    
    [
        pg.Text
        (
            "",
            size=(50, 16),
            pad=(0, 0),
            key="-MISSING-"
        )
    ],
    [
        pg.Button
        (
            "Open File", 
            size=(10, 2),
        ),
        pg.Button
        (
            "Open File Location", 
            pad=(60, 0),
            size=(10, 2),
        ),
        pg.Button
        (
            "Open All Files", 
            size=(10, 2),
        )
    ]
]

layout = [    
    [        
        pg.Column(file_list_column, expand_y=True),         
        pg.VSeperator(),        
        pg.Column(file_viewer_column, expand_y=True)    
    ]
]


# Home directory
homeFolder = Path.home()

# OneDrive directory
oneDriveFolder = os.path.join(homeFolder, 'OneDrive')

# Trace directory
traceFolder = os.path.join(oneDriveFolder, 'Traceability')


# Trace function to search data excel file
def trace(number):
    # Traceability Data excel file directory
    df = pd.read_excel(os.path.join(traceFolder, 'Traceability Data.xlsx'))

    # Dictionary to store information
    temp = {}

    # Get Work Order number
    WO = number                                             

    # Get row index of work order
    if WO in df["Work Order"].values:
        rowIndex = df[df['Work Order']==WO].index.item()
        temp['Work Order'] = f'WO{WO}'
    else:
        temp['Work Order'] = 'Not Found'
        return temp

    # Get DO using rowIndex and update dictionary
    if pd.isnull(df.loc[rowIndex, 'DO']):
        temp['DO'] = 'Not Found'
    else:
        DO = df.loc[rowIndex, 'DO']
        temp['DO'] = int(DO)

    # Get PO using rowIndex and update dictionary
    if pd.isnull(df.loc[rowIndex, 'PO']):
        temp['PO'] = 'Not Found'
    else:
        PO = df.loc[rowIndex, 'PO']
        temp['PO'] = int(PO)

    # Get IPQC using rowIndex and update dictionary
    if pd.isnull(df.loc[rowIndex, 'IPQC']):
        temp['IPQC'] = 'Not Found'
    else:
        IPQC = df.loc[rowIndex, 'IPQC']
        temp['IPQC'] = f'WO{int(IPQC)}' 

    return temp


# Create Window
window = pg.Window("Work Order Trace", layout, finalize=True, size=(800, 420))
window['-INPUT-'].bind("<Return>", "_Enter")


# Event loop
while True:    
    event, values = window.read()  
    if event == pg.WIN_CLOSED:        
        break    

    # Invoke trace function when Search button is clicked or Enter is pressed
    if event == "Search" or event == "-INPUT-" + "_Enter":
        try:
            window["-MISSING-"].update('')
            temp = trace(int(values["-INPUT-"]))
            fileNames = []                              # File names and folders to be displayed
            filePaths = []                              # Full paths of files

            # loop through the dictionary and search directory
            for k, v in temp.items():
                p = Path(traceFolder, k)
                if v == 'Not Found':
                    window["-MISSING-"].update(window["-MISSING-"].get() + '\n' + f'{k} is {v}')
                for root, dirs, files in os.walk(p):
                    for file in files:
                        if re.search(v, file):
                            fullFile = os.path.join(root, file)
                            filePaths.append(fullFile)
                            fileNames.append('\\'.join(fullFile.split('\\')[-2:])) 
                        
            window["-FILE_LIST-"].update(fileNames)
            window["-DESC-"].update("This list does not contain Customer PO, MRF, Mass Balance and CCP")
        except:
            files = []
          
    # Show path directory of the selected file
    if event == "-FILE_LIST-" and len(values["-FILE_LIST-"]) > 0:        
        file_selection = values["-FILE_LIST-"][0]  
        for file in filePaths:
            if file.endswith(file_selection):
                window["-TOUT-"].update(file)
 
    # Open the selected file
    if event == "Open File" and len(values["-FILE_LIST-"]) > 0:
        for file in filePaths:
            if file.endswith(file_selection):
                os.startfile(file)

    # Open the selected file's directory
    if event == "Open File Location" and len(values["-FILE_LIST-"]) > 0:
        for file in filePaths:
            if file.endswith(file_selection):
                os.startfile('\\'.join(file.split('\\')[:-1]))

    # Open all the files in the directory
    if event == "Open All Files":
        try:
            for file in filePaths:
                os.startfile(file)
        except:
            files = []


# Close window
window.close()