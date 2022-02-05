
import pandas as pd
from IPython.display import display
import os


# requires installation  pandas, openpyxl, IPython (for instance via 
# pip install pandas openpyxl IPython
listOfFiles = [[],[]] # empty array, will contain to subarrays, one with the csv file names, one with the xlsx file names



for files in os.listdir('csv/'): # listdir lists all files in a given directory
    if files.endswith('.csv'): # ignores non .csv files
        listOfFiles[0].append(files) # appends the filename in the first subarray ( count starts with 0)
    else:
        continue
for files in os.listdir('xlsx'):
    if files.endswith('.xlsx'):
        listOfFiles[1].append(files)
    else:
        continue


print(listOfFiles)

def logic(index): # returns true for even numbers, false for odd numbers
    if index % 2 == 0:
       return True
    return False

dataframe =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it

## Comment this in if you want to parse .csv files
#os.chdir('csv/') # Changes working directory to the csv subfolder

# for x in listOfFiles[0]: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe
#     i = pd.read_csv(
#         x,
#         header=6,
#         usecols = [0, 1, 7, 10], # column number, first is 0. 0 = Title, 7 = Online_ISSN, 10 = Reporting_Period_Total
#         #skiprows= lambda x: logic(x), # calls the logic function declared above, skips uneven row numbers (the rows with total requests)
#         )
#     #display(i)
#     dataframe = dataframe.append(i, ignore_index=True) # appends adds each csv content to end of the dataframe, ignore_index is there so he doesnt also print the original row number

os.chdir('xlsx') #Changes working directory to the xlsx subfolder



for x in listOfFiles[1]: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe
    i = pd.read_excel(
        x,
        header=13,
        usecols = [0, 1, 7, 9, 10], # column number, first is 0. 0 = Title, 1 = Publisher, 7 = Online_ISSN, 9 = Metric_Type, 10 = Reporting_Period_Total
        #skiprows= lambda x: logic(x), # calls the logic function declared above, skips uneven row numbers (the rows with total requests)
        )
    dataframe = dataframe.append(i, ignore_index=True) # appends adds each csv content to end of the dataframe, ignore_index is there so he doesnt also print the original row number


os.chdir('../') # changes working directory to main folder 

display(dataframe) # optional, shows the new table in the console
dataframe = dataframe.loc[dataframe['Metric_Type'] == "Unique_Item_Requests"]
display(dataframe) # optional, shows the new table in the console
dataframe.to_csv("masterlist.csv", index=False) # writes the dataframe to a new file
        


# print(listOfFiles[0])
# df=pd.read_csv(
# listOfFiles[0], 
# header=6,
# usecols = [0, 7, 9, 10],
# skiprows= lambda x: logic(x),
# )#["Title", "Online_ISSN",])# "Reporting_Period_Total"]) # only for unique item request
# display(df)

#df.to_csv("dataframe2.csv", index=False)