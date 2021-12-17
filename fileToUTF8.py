from typing import cast
import pandas as pd
from os import listdir
from os.path import isfile, join

input_path = 'Excel/'
unclean_path = 'CSV-UTF8-Unclean/'
clean_path = 'CSV-UTF8-Clean/'
json_path = 'JSON/'

# Getting all excel files inside Excel folder and store it in onlyfiles
onlyfiles = [f for  f in listdir(input_path) if isfile(join(input_path,f))]

for file in onlyfiles:
    # Reading excel files from Excel folder and removing unwanted columns
    data = pd.read_excel(input_path+file)

    # Converting excel file into csv and encoding to utf-8 format
    data.to_csv(unclean_path+file.split('.')[0]+'.csv',index=False,encoding='utf-8')

# Getting all Unclean files from CSV-UTF8-Unclean folder
uncleanfiles = [f for f in listdir(unclean_path) if isfile(join(unclean_path,f))]

for file in uncleanfiles:
    # Reading csv-utf8 unclean files
    data = pd.read_csv(unclean_path+file)
    
    # Removing Unwanted columns in data
    try:
        data.drop(columns=['Unnamed: 15','Unnamed: 16'])
    except:
        print('No columns found to drop in file ',file)

    # Saving the csv file to CSV-UTF8-Clean folder
    data.to_csv(clean_path+file,index=False)

# Getting all files inside CSV-UTF8-Clean folder
cleanfiles = [f for f in listdir(clean_path) if isfile(join(clean_path,f))]

for file in cleanfiles:
    # Reading clean file 
    data = pd.read_csv(clean_path+file)

    # Saving the csv file in json format and store inside JSON folder
    data.to_json(json_path+file.split('.')[0]+'.json',orient='records')




