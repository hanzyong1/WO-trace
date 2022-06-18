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
            size=(10, 1),
        )
    ], 
    [        
        pg.Listbox(            
            values=[],            
            enable_events=True,            
            size=(50, 30),            
            key="-FILE_LIST-",
        )    
    ]
]

file_viewer_column = [    
    [
        pg.Text
        (
            "Choose a file from the list", 
            size=(50, 1),
            pad=(10, 20)
        )
    ],    
    [
        pg.Text
        (
            "File name: ", 
            size=(50, 2), 
            pad=(10, 20),
            key="-TOUT-"
        )
    ],    
    [
        pg.Button
        (
            "Open File", 
            pad=(10, 20),
            size=(10, 2),
        )
    ],
    [
        pg.Button
        (
            "Open File Location", 
            pad=(10, 20),
            size=(10, 2),
        )
    ],
    [
        pg.Button
        (
            "Open All Files", 
            pad=(10, 20),
            size=(10, 2),
        )
    ],
]

layout = [    
    [        
        pg.Column(file_list_column),        
        pg.VSeperator(),        
        pg.Column(file_viewer_column)    
    ]
]


# Home directory
homeFolder = Path.home()

# OneDrive directory
oneDriveFolder = os.path.join(homeFolder, 'OneDrive')


# Trace function to search data excel file
def trace(number):
    # Traceability Data excel file directory
    df = pd.read_excel(os.path.join(oneDriveFolder, 'Traceability Data', 'Traceability Data.xlsx'))

    # Dictionary to store information
    temp = {}

    # Get Work Order number
    WO = number                                             
    temp['Work Order'] = f'WO{WO}'

    # Get row index of work order
    rowIndex = df[df['Work Order']==WO].index.item()

    # Get DO using rowIndex and update dictionary
    DO = df.loc[rowIndex, 'DO']
    temp['DO'] = DO

    # Get PO using rowIndex and update dictionary
    PO = df.loc[rowIndex, 'PO']
    temp['PO'] = PO

    # Get Batch Number using rowIndex and update dictionary
    # batchNumber = df.loc[rowIndex, 'Batch Number']
    # temp['Batch Number'] = batchNumber

    # Get Roll ID using rowIndex and update dictionary
    # rollID = df.loc[rowIndex, 'Roll ID']
    # temp['Roll ID'] = rollID
    return temp


# Create Window
window = pg.Window("File Viewer", layout, finalize=True)
window['-INPUT-'].bind("<Return>", "_Enter")


# Event loop
while True:    
    event, values = window.read()  
    if event == pg.WIN_CLOSED:        
        break    

    # Invoke trace function when Search button is clicked or Enter is pressed
    if event == "Search" or event == "-INPUT-" + "_Enter":
        try:
            temp = trace(int(values["-INPUT-"]))
            file_names = []
        
            # loop through the dictionary and search directory
            for k, v in temp.items():
                p = Path(oneDriveFolder, k)
                for root, dirs, files in os.walk(p):
                    for file in files:
                        if file == f"{v}.pdf":
                            file_names.append(file)       
            window["-FILE_LIST-"].update(file_names)
        except:
            files = []
          
    # Show path directory of the selected file
    if event == "-FILE_LIST-" and len(values["-FILE_LIST-"]) > 0:        
        file_selection = values["-FILE_LIST-"][0]  
        for k, v in temp.items():
            if file_selection == f"{v}.pdf":
                window["-TOUT-"].update(Path(oneDriveFolder, k, file_selection))
 
    # Open the selected file
    if event == "Open File" and len(values["-FILE_LIST-"]) > 0:
        for k, v in temp.items():
            if file_selection == f"{v}.pdf":
                os.startfile(os.path.join(oneDriveFolder, k, file_selection))

    # Open the selected file's directory
    if event == "Open File Location" and len(values["-FILE_LIST-"]) > 0:
        for k, v in temp.items():
            if file_selection == f"{v}.pdf":
                os.startfile(os.path.join(oneDriveFolder, k))

    # Open all the files in the directory
    if event == "Open All Files":
        for k, v in temp.items():
            os.startfile(os.path.join(oneDriveFolder, k, f"{v}.pdf"))   


# Close window
window.close()
