import pandas as pd
from IPython.display import display
import os


errorMessages= []
# First we need arrays containing a list of all the filenames in the subdirectories
listOfSourceFiles_C_4 = [] 
listOfSourceFiles_C_5 = [] # empty array, will contain to subarrays, one with the csv file names, one with the xlsx file names
listOfExcludeFiles = []
listOfSingleJournalFiles = []


for files in os.listdir('C_4'):
    if files.endswith('.xlsx'):
        listOfSourceFiles_C_4.append(files)
    else:
        continue



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

for files in os.listdir('single_journals'):
    if files.endswith('.xlsx'):
        listOfSingleJournalFiles.append(files)
    else:
        continue

print("xlsx-Dateien im C_4 Verzeichnis: \n", listOfSourceFiles_C_4, "\n")
print("xlsx-Dateien im C_5 Verzeichnis: \n", listOfSourceFiles_C_5, "\n")
print("xlsx-Dateien im exclude Verzeichnis: \n", listOfExcludeFiles, "\n")
print("xlsx-Dateien im single_journal Verzeichnis: \n", listOfSingleJournalFiles, "\n")



dataframe_titles_C_4 =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it
dataframe_titles_C_5 =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to 
dataframe_exclude =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it
dataframe_single_journals = pd.DataFrame()




os.chdir('C_4') # Opens C_4 directory
for x in listOfSourceFiles_C_4: #loops through the list of files in the C_4 directory, and appends each to dataframe_titles_C_4
    try:
        i = pd.read_excel(
            x,
            header=7,      # Pandas starts the count with 0, unlike Excel which starts at 1 
            skiprows=[8],
            usecols = ["Title", "Publisher", "Print ISSN", "Online ISSN", "Reporting_Period_Total"], 
            ).assign(C_4_Reportname=x)
        dataframe_titles_C_4 = dataframe_titles_C_4.append(i, ignore_index=True) # Jeder Schleifendurchlauf erweitert die Tabelle um die neuen Werte
    except Exception as e: errorMessages.append(("Error in C_4/%s:" %x, e ))


os.chdir('../C_5')
for x in listOfSourceFiles_C_5: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe_titles_C_5
    try:
        i = pd.read_excel(
            x,
            header=14,
            usecols = ["Title", "Publisher", "Print_ISSN","Online_ISSN", "Metric_Type", "Reporting_Period_Total"], 
            ).assign(C_5_Reportname=x)
        dataframe_titles_C_5 = dataframe_titles_C_5.append(i, ignore_index=True) # appends adds each csv content to end of the dataframe_titles_C_5, ignore_index is there so he doesnt also print the original row number
    except Exception as e: errorMessages.append(("Error in C_5/%s:" %x, e ))


#display(dataframe_titles_C_5) # optional, shows the new table in the console
dataframe_titles_C_5 = dataframe_titles_C_5.loc[dataframe_titles_C_5['Metric_Type'] == "Unique_Item_Requests"] #removes all rows in which Metric_Type isnt Unique_Item_Requests
#print("Only unique_item_requests \n")
#display(dataframe_titles_C_5) # optional, shows the new table in the console

os.chdir('../exclude')

for x in listOfExcludeFiles: #loops through the list of files in the csv and xlsx directory, and appends each to dataframe_titles_C_5
    try:
        i = pd.read_excel(
            x,
            header=0,
            usecols = ["Title"], 
            )
        dataframe_exclude = dataframe_exclude.append(i, ignore_index=True) # appends adds each xlsx files content to end of the dataframe_titles_C_5, ignore_index is there so he doesnt also print the original row number
    except Exception as e:  errorMessages.append(("Error in exclude/%s:" %x, e ))


exludeList = dataframe_exclude["Title"].tolist() #isin()



master = pd.DataFrame()

dataframe_titles_C_4.columns=["Title", "Publisher", "Print_ISSN","Online_ISSN", "Reporting_Period_Total"]
#dataframe_titles_C_4 = dataframe_titles_C_4.append(dataframe_titles_C_5)

master = master.append(dataframe_titles_C_4)
master = master.append(dataframe_titles_C_5)


print("\n Ungefilterte Mastertabelle: \n",master, "\n")
master.to_csv("../master_unfiltered.csv", index=False) # Enthält noch die Medizintitel und irrelevante Zeilen ohne Reporting_Period_Total sind noch enthalten
master = master.loc[~master['Title'].isin(exludeList)] # Removes the titles found in files saved in the "exclude" directory



emptyReporting_Period_Total_values = master.loc[master['Reporting_Period_Total'].isna(), :] # List containing the rows with empty Reporting_Period_Total

master = master.loc[~master['Reporting_Period_Total'].isna()] # Removes rows in which Reporting_Period_Total is empty
print("Finale Masterliste: \n", master, "\n")
master.to_csv("../master_without_excluded_titles.csv", index=False)




print("Deleted Rows with empty Reporting_Period_Total: \n", emptyReporting_Period_Total_values)


#Parsing List of Single Journals

os.chdir('../single_journals') # Opens single_journal directory
for x in listOfSingleJournalFiles: #loops through the list of files in the single_journal directory, and appends each to dataframe_titles_C_4
    try:
        i = pd.read_excel(
            x,
            usecols = ["Title", "Verlag", "ISSN", "Preis 2021"],
            )
        dataframe_single_journals = dataframe_single_journals.append(i, ignore_index=True) # Jeder Schleifendurchlauf erweitert die Tabelle um die neuen Werte
    except Exception as e: errorMessages.append(("Error in single_journals/%s:" %x, e ))

os.chdir('../') # changes working directory to main folder 


df_single_journals_with_ISSN = pd.DataFrame()
df_single_journals_without_ISSN = pd.DataFrame()
df_single_journals_with_ISSN = dataframe_single_journals.loc[~dataframe_single_journals['ISSN'].isna(), :] # List containing the rows with ISSN
df_single_journals_without_ISSN = dataframe_single_journals.loc[dataframe_single_journals['ISSN'].isna(), :] # List containing the rows without ISSN
df_single_journals_without_ISSN.to_csv("Einzelkaufslisteneinträge ohne ISSN.csv", index=False)


df_print_ISSN = pd.DataFrame()
df_print_ISSN = df_single_journals_with_ISSN.merge(right=master, how="inner", left_on=["ISSN"], right_on=["Print_ISSN", ] )

df_single_journals_with_ISSN = df_single_journals_with_ISSN.merge(right=master, how="inner", left_on=["ISSN"], right_on=["Online_ISSN", ] )

df_single_journals_with_ISSN = df_single_journals_with_ISSN.loc[~df_single_journals_with_ISSN['Preis 2021'].isna()]
df_single_journals_with_ISSN['Preis 2021/Reporting_Period_Total'] = df_single_journals_with_ISSN['Preis 2021']/df_single_journals_with_ISSN['Reporting_Period_Total']
df_single_journals_with_ISSN['Preis 2021/Reporting_Period_Total'] = df_single_journals_with_ISSN['Preis 2021/Reporting_Period_Total'].round(2)
print("\nMatched Online_ISSN and calculated Preis 2021/Reporting_Period_Total\n", df_single_journals_with_ISSN)
df_single_journals_with_ISSN.to_csv("ISSN_merged.csv", index=False)




print("print_issn matching \n",df_print_ISSN)
df_print_ISSN.to_csv("print.csv", index=False)
#Print Errors, if any
if (len(errorMessages) > 0 and len(errorMessages) < 2):
    print("\033[0;31m", "An Error occured: \n", errorMessages, "\033[0m")
if (len(errorMessages) > 1 ):
    print("\033[0;31m", "Multiple Errors occured: \n", errorMessages, "\033[0m")


#print(e)
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
