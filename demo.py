import PySimpleGUI as pg
import os

# Step 1: Set theme
pg.theme("SystemDefault1")

# Step 2: Create layout
file_list_column = [    
    [        
        pg.Text("Folder"),        
        pg.In(size=(30, 1), 
        enable_events=True, 
        key="-FOLDER-"),        
        pg.FolderBrowse(),    
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

# Step 3: Create Window
window = pg.Window("File Viewer", layout)

# Step 4: Event loop
folder_location = ""
all_files = []                  # empty list to store all the files of the directory

while True:    
    event, values = window.read()  
    if event == pg.WIN_CLOSED:        
        break    
    if event == "-FOLDER-":        
        folder_location = values["-FOLDER-"]        
        try:            
            files = os.listdir(folder_location)        
        except:            
            files = []                
            
        file_names = [            
            file for file in files            
            if os.path.isfile(os.path.join(folder_location, file)) and 
            file.lower().endswith((".pdf", ".csv"))             # only certain file types
        ]        
        window["-FILE_LIST-"].update(file_names)
        all_files = file_names                                 # list of all the files in the directory
    
    if event == "-FILE_LIST-" and len(values["-FILE_LIST-"]) > 0:        
        file_selection = values["-FILE_LIST-"][0]     
        window["-TOUT-"].update(os.path.join(folder_location, file_selection)) 
 
    # Open the selected file
    if event == "Open File" and len(values["-FILE_LIST-"]) > 0:
        os.startfile(os.path.join(folder_location, file_selection))

    # Open the selected file's directory
    if event == "Open File Location" and len(values["-FILE_LIST-"]) > 0:
        os.startfile(os.path.join(folder_location))

    # Open all the files in the directory
    if event == "Open All Files":
        for file in all_files:
            os.startfile(os.path.join(folder_location, file))

# Step 5: Close window
window.close()
