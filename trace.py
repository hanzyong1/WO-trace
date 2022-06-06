import pandas as pd
from pathlib import Path
import os

# read database
df = pd.read_excel('test.xlsx')                 # testing purposes

# dictionary to store information
temp = {}

# get Work Order number
# WO = int(input('Please enter a number '))           # for user to input Work Order
WO = 1                                              # testing purposes
temp['Work Order'] = f'WO{WO}'

# get row index of work order
rowIndex = df[df['Work Order']==WO].index.item()

# get Batch Number using rowIndex and update dictionary
batchNumber = df.loc[rowIndex, 'Batch Number']
temp['Batch Number'] = batchNumber

# get Roll ID using rowIndex and update dictionary
rollID = df.loc[rowIndex, 'Roll ID']
temp['Roll ID'] = rollID

# get DO using rowIndex and update dictionary
DO = df.loc[rowIndex, 'DO']
temp['DO'] = DO

# home directory
homeFolder = Path.home()

# OneDrive directory
oneDriveFolder = Path('OneDrive')

# testing directory
testFolder = Path('python', 'hexachase')

# loop through the dictionary and search directory
for k, v in temp.items():
    searchFolder = Path(testFolder, k)                  # for testing purposes
    # searchFolder = Path(oneDriveFolder, k)              # to search OneDrive
    p = homeFolder/searchFolder
    for root, dirs, files in os.walk(p):
        for file in files:
            if file == f'{v}.pdf':
                os.startfile(os.path.join(p, file))