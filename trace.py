import PySimpleGUI as pg
import os
import pandas as pd
from pathlib import Path


# Set theme
pg.theme("SystemDefault1")


# Create layout
file_list_column = [      
    [
        pg.Input(
        "",
        size=(30, 1), 
        enable_events=True, 
        key="-INPUT-"
        ),
        pg.Button
        (
            "Search", 
            pad=(10, 10),
            size=(8, 1),
        )
    ], 
    [        
        pg.Listbox(            
            values=[],            
            enable_events=True,            
            size=(40, 20),            
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
            pad=(0 ,0),
            key="-DESC-"
        )
    ],    
    [
        pg.Text
        (
            "File path: ", 
            size=(50, 1), 
            pad=(0, 10),
            key="-TOUT-"
        )
    ],    
    [
        pg.Text
        (
            "",
            size=(50, 13),
            pad=(0, 10),
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
            pad=(50, 0),
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
        pg.Column(file_list_column),        
        pg.VSeperator(),        
        pg.Column(file_viewer_column, expand_y=True, pad=((5, 0), (15, 0)))    
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

    return temp


# Create Window
window = pg.Window("Work Order Trace", layout, finalize=True, size=(750, 400))
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
            file_names = []
        
            # loop through the dictionary and search directory
            for k, v in temp.items():
                p = Path(traceFolder, k)
                if v == 'Not Found':
                    window["-MISSING-"].update(window["-MISSING-"].get() + '\n' + f'{k} is {v}')
                for root, dirs, files in os.walk(p):
                    for file in files:
                        if file == f"{v}.pdf":
                            file_names.append(file)  
                            
            window["-FILE_LIST-"].update(file_names)
            window["-DESC-"].update("This list does not contain Customer PO, MRF and Mass Balance")
        except:
            files = []
          
    # Show path directory of the selected file
    if event == "-FILE_LIST-" and len(values["-FILE_LIST-"]) > 0:        
        file_selection = values["-FILE_LIST-"][0]  
        for k, v in temp.items():
            if file_selection == f"{v}.pdf":
                window["-TOUT-"].update(Path(traceFolder, k, file_selection))
 
    # Open the selected file
    if event == "Open File" and len(values["-FILE_LIST-"]) > 0:
        for k, v in temp.items():
            if file_selection == f"{v}.pdf":
                os.startfile(os.path.join(traceFolder, k, file_selection))

    # Open the selected file's directory
    if event == "Open File Location" and len(values["-FILE_LIST-"]) > 0:
        for k, v in temp.items():
            if file_selection == f"{v}.pdf":
                os.startfile(os.path.join(traceFolder, k))

    # Open all the files in the directory
    if event == "Open All Files":
        try:
            for k, v in temp.items():
                os.startfile(os.path.join(traceFolder, k, f"{v}.pdf"))   
        except:
            files = []


# Close window
window.close()
