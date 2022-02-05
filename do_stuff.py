
import pandas as pd
from IPython.display import display
import os


listOfSourceFiles_C_4 = [] 
listOfSourceFiles_C_5 = [] # empty array, will contain to subarrays, one with the csv file names, one with the xlsx file names
listOfExcludeFiles = []


for files in os.listdir('C_4'):
    if files.endswith('.xlsx'):
        listOfSourceFiles_C_4.append(files)
    else:
        continue

print("C_4 Files: \n", listOfSourceFiles_C_4)

for files in os.listdir('C_5'):
    if files.endswith('.xlsx'):
        listOfSourceFiles_C_5.append(files)
    else:
        continue

for files in os.listdir('exclude'):
    if files.endswith('.xlsx'):
        listOfExcludeFiles.append(files)
    else:
        continue


print(listOfSourceFiles_C_5)
print(listOfExcludeFiles)



dataframe_titles_C_5 =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to 
dataframe_titles_C_4 =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it
dataframe_exclude =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it




os.chdir('C_4')
for x in listOfSourceFiles_C_4: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe_titles_C_5
    print("Start parsing ",x)
    i = pd.read_excel(
        x,
        header=7,      
        skiprows=[8],
        usecols = ["Title", "Publisher", "Online ISSN", "Reporting_Period_Total"], 
        )
    dataframe_titles_C_4 = dataframe_titles_C_4.append(i, ignore_index=True) # appends adds each csv content to end of the dataframe_titles_C_5, ignore_index is there so he doesnt also print the original row number
    print("Successfully parsed ",x)
print("C_4 Titles \n")
display(dataframe_titles_C_4)
os.chdir('../C_5')
for x in listOfSourceFiles_C_5: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe_titles_C_5
    print("Start parsing ",x)
    i = pd.read_excel(
        x,
        header=14,
        usecols = ["Title", "Publisher", "Online_ISSN", "Metric_Type", "Reporting_Period_Total"], 
       
        )
    dataframe_titles_C_5 = dataframe_titles_C_5.append(i, ignore_index=True) # appends adds each csv content to end of the dataframe_titles_C_5, ignore_index is there so he doesnt also print the original row number
    print("Successfully parsed ",x)



#display(dataframe_titles_C_5) # optional, shows the new table in the console
dataframe_titles_C_5 = dataframe_titles_C_5.loc[dataframe_titles_C_5['Metric_Type'] == "Unique_Item_Requests"] #removes all rows in which Metric_Type isnt Unique_Item_Requests
#print("Only unique_item_requests \n")
#display(dataframe_titles_C_5) # optional, shows the new table in the console

os.chdir('../exclude')
for x in listOfExcludeFiles: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe_titles_C_5
    print("Start parsing ",x)
    i = pd.read_excel(
        x,
        header=0,
        usecols = ["Title"], 
        )
    dataframe_exclude = dataframe_exclude.append(i, ignore_index=True) # appends adds each xlsx files content to end of the dataframe_titles_C_5, ignore_index is there so he doesnt also print the original row number
    print("Successfully parsed ",x)

os.chdir('../') # changes working directory to main folder 

print("exclude df: \n")
display(dataframe_exclude)

exludeList = dataframe_exclude["Title"].tolist() #isin()



master = pd.DataFrame(columns=["Title", "Publisher", "Online_ISSN", "Metric_Type", "Reporting_Period_Total"])

master = dataframe_titles_C_4
master = master.append(dataframe_titles_C_5)
display(master)
master.to_csv("master_unfiltered.csv", index=False) # Unfiltered mastertable
master = master.loc[~master['Title'].isin(exludeList)] # Removes the titles found in files saved in the "exclude" directory


display(master)
emptyReporting_Period_Total_values = master.loc[master['Reporting_Period_Total'].isnull(), :] # List containing the rows with empty Reporting_Period_Total


master = master.loc[~master['Reporting_Period_Total'].isna()] # Removes rows in which Reporting_Period_Total is empty
display(master)

master.to_csv("master_without_excluded_titles.csv", index=False)

print("Deleted Rows with empty Reporting_Period_Total: \n", emptyReporting_Period_Total_values)


# Below is old unused code

# print(listOfSourceFiles_C_5[0])
# df=pd.read_csv(
# listOfSourceFiles_C_5[0], 
# header=6,
# usecols = [0, 7, 9, 10],
# skiprows= lambda x: logic(x),
# )#["Title", "Online_ISSN",])# "Reporting_Period_Total"]) # only for unique item request
# display(df)

#df.to_csv("dataframe_titles_C_52.csv", index=False)

# Comment in if you have csv source files
# for files in os.listdir('csv/'): # listdir lists all files in a given directory
#     if files.endswith('.csv'): # ignores non .csv files
#         listOfSourceFiles_C_5[0].append(files) # appends the filename in the first subarray ( count starts with 0)
#     else:
#         continue