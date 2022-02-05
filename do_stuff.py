
import pandas as pd
from IPython.display import display
import os



listOfSourceFiles = [[],[]] # empty array, will contain to subarrays, one with the csv file names, one with the xlsx file names
listOfExcludeFiles = []

# Comment in if you have csv source files
# for files in os.listdir('csv/'): # listdir lists all files in a given directory
#     if files.endswith('.csv'): # ignores non .csv files
#         listOfSourceFiles[0].append(files) # appends the filename in the first subarray ( count starts with 0)
#     else:
#         continue
for files in os.listdir('xlsx'):
    if files.endswith('.xlsx'):
        listOfSourceFiles[1].append(files)
    else:
        continue

for files in os.listdir('exclude'):
    if files.endswith('.xlsx'):
        listOfExcludeFiles.append(files)
    else:
        continue


print(listOfSourceFiles)
print(listOfExcludeFiles)

def logic(index): # returns true for even numbers, false for odd numbers
    if index % 2 == 0:
       return True
    return False

dataframe_titles_C_5 =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to 
dataframe_titles_C_4 =pd.DataFrame(columns=['Titles', 'Publisher', 'Online_ISSN',]) #Empty object, the loop below adds the Columns and Rows to it
dataframe_exclude =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it

## Comment this in if you want to parse .csv files
#os.chdir('csv/') # Changes working directory to the csv subfolder

# for x in listOfSourceFiles[0]: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe_titles_C_5
#     i = pd.read_csv(
#         x,
#         header=6,
#         usecols = [0, 1, 7, 10], # column number, first is 0. 0 = Title, 7 = Online_ISSN, 10 = Reporting_Period_Total
#         #skiprows= lambda x: logic(x), # calls the logic function declared above, skips uneven row numbers (the rows with total requests)
#         )
#     #display(i)
#     dataframe_titles_C_5 = dataframe_titles_C_5.append(i, ignore_index=True) # appends adds each csv content to end of the dataframe_titles_C_5, ignore_index is there so he doesnt also print the original row number

os.chdir('xlsx') #Changes working directory to the xlsx subfolder



for x in listOfSourceFiles[1]: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe_titles_C_5
    i = pd.read_excel(
        x,
        header=14,
        usecols = ["Title", "Publisher", "Online_ISSN", "Metric_Type", "Reporting_Period_Total"], # column number, first is 0. 0 = Title, 1 = Publisher, 7 = Online_ISSN, 9 = Metric_Type, 10 = Reporting_Period_Total
        #skiprows= lambda x: logic(x), # calls the logic function declared above, skips uneven row numbers (the rows with total requests)
        )
    dataframe_titles_C_5 = dataframe_titles_C_5.append(i, ignore_index=True) # appends adds each csv content to end of the dataframe_titles_C_5, ignore_index is there so he doesnt also print the original row number




display(dataframe_titles_C_5) # optional, shows the new table in the console
dataframe_titles_C_5 = dataframe_titles_C_5.loc[dataframe_titles_C_5['Metric_Type'] == "Unique_Item_Requests"] #removes all rows in which Metric_Type isnt Unique_Item_Requests
print("Only unique_item_requests \n")
display(dataframe_titles_C_5) # optional, shows the new table in the console

os.chdir('../exclude')
for x in listOfExcludeFiles: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe_titles_C_5
    i = pd.read_excel(
        x,
        header=0,
        usecols = ["Title"], # column number, first is 0. 0 = Title, 1 = ISSN
        #skiprows=0,
        #skiprows= lambda x: logic(x), # calls the logic function declared above, skips uneven row numbers (the rows with total requests)
        )
    dataframe_exclude = dataframe_exclude.append(i, ignore_index=True) # appends adds each csv content to end of the dataframe_titles_C_5, ignore_index is there so he doesnt also print the original row number
os.chdir('../') # changes working directory to main folder 
#dataframe_titles_C_5 = dataframe_titles_C_5.loc[dataframe_titles_C_5['Metric_Type'] == "Unique_Item_Requests"]
print("exclude df: \n")
display(dataframe_exclude)
dataframe_titles_C_5.to_csv("masterlist_including_medicine.csv", index=False) # writes the dataframe_titles_C_5 to a new file
#dataframe_exclude = dataframe_exclude.loc[dataframe_exclude['Metric_Type'] == "Unique_Item_Requests"]
#dataframe_exclude = dataframe_exclude.isin()
exludeList = dataframe_exclude["Title"].tolist() #isin()
dataframe_titles_C_5 = dataframe_titles_C_5.loc[~dataframe_titles_C_5['Title'].isin(exludeList)]
display(dataframe_titles_C_5)

dataframe_titles_C_5.to_csv("masterlist_without_medicine.csv", index=False) # writes the dataframe_titles_C_5 to a new file



# print(listOfSourceFiles[0])
# df=pd.read_csv(
# listOfSourceFiles[0], 
# header=6,
# usecols = [0, 7, 9, 10],
# skiprows= lambda x: logic(x),
# )#["Title", "Online_ISSN",])# "Reporting_Period_Total"]) # only for unique item request
# display(df)

#df.to_csv("dataframe_titles_C_52.csv", index=False)