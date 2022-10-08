import pandas as pd
from IPython.display import display
import os
# openpyxl too!

#Function Declarations

# removes rows in which the specified column is empty
def remove_empty_rows (dataframe, column):
    print(column)
    dataframe = dataframe.loc[~dataframe[column].isna()]

    return dataframe

    
# parses xlsx documents, directory is a mandatory parameter and is defined relative to the directory the script is running in, the other ones are optional
def parse_xlsx (directory,  skiprows=None, header= 0,  usecolumns = None, assign = False ):
    listOfFiles = []
    dataframe = pd.DataFrame()

    for files in os.listdir(directory):
        if files.endswith('.xlsx'):
            listOfFiles.append(files)
        else:
            continue
    print(listOfFiles)
    os.chdir(directory)

    for File in listOfFiles: 
        #try:
            i = pd.read_excel(
                File,
                header=header,      # Pandas starts the count with 0, unlike Excel which starts at 1 
                skiprows=skiprows,
                usecols = usecolumns, 
                )
            if assign == True:
                i=i.assign(Sourcefilename=File) 
            dataframe = dataframe.append(i, ignore_index=True) # Jeder Schleifendurchlauf erweitert die Tabelle um die neuen Werte
            
        #except Exception as e: errorMessages.append(("Error in ", directory, "/", File, ": ",  e ))
        #print("test")

    os.chdir("../")
    
    return dataframe



# First we need arrays containing a list of all the filenames in the subdirectories
# listOfSourceFiles_C_4 = [] 
# listOfSourceFiles_C_5 = [] # empty array, will contain to subarrays, one with the csv file names, one with the xlsx file names
# listOfExcludeFiles = []
# listOfSingleJournalFiles = []


# for files in os.listdir('C_4'):
#     if files.endswith('.xlsx'):
#         listOfSourceFiles_C_4.append(files)
#     else:
#         continue



# # for files in os.listdir('C_5'):
#     if files.endswith('.xlsx'):
#         listOfSourceFiles_C_5.append(files)
#     else:
#         continue

# for files in os.listdir('exclude'):
#     if files.endswith('.xlsx'):
#         listOfExcludeFiles.append(files)
#     else:
#         continue

# for files in os.listdir('single_journals'):
#     if files.endswith('.xlsx'):
#         listOfSingleJournalFiles.append(files)
#     else:
#         continue

# print("xlsx-Dateien im C_4 Verzeichnis: \n", listOfSourceFiles_C_4, "\n")
# print("xlsx-Dateien im C_5 Verzeichnis: \n", listOfSourceFiles_C_5, "\n")
# print("xlsx-Dateien im exclude Verzeichnis: \n", listOfExcludeFiles, "\n")
# print("xlsx-Dateien im single_journal Verzeichnis: \n", listOfSingleJournalFiles, "\n")


errorMessages= []
dataframe_titles_C_4 =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it
dataframe_titles_C_5 =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to 
dataframe_exclude =pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it
dataframe_single_journals = pd.DataFrame() #Empty object, the loop below adds the Columns and Rows to it

dataframe_titles_C_4 = parse_xlsx("C_4", [8], 7, ["Title", "Publisher", "Platform", "Print ISSN", "Online ISSN", "Reporting_Period_Total"], assign=True)

#print(dataframe_titles_C_4)
dataframe_titles_C_5 = parse_xlsx("C_5", [], 14, ["Title", "Publisher", "Platform", "Print_ISSN","Online_ISSN", "Metric_Type", "Reporting_Period_Total"], assign=True)

#display(dataframe_titles_C_5) # optional, shows the new table in the console

#removes all rows in which Metric_Type isnt Unique_Item_Request
dataframe_titles_C_5 = dataframe_titles_C_5.loc[dataframe_titles_C_5['Metric_Type'] == "Unique_Item_Requests"] 


dataframe_exclude = parse_xlsx(directory="exclude")

print(dataframe_exclude)


master = pd.DataFrame()

# Columns are differently named in C_4, so I have to rename them to be the same as in the C_5 Standard
dataframe_titles_C_4.columns=["Title", "Publisher", "Platform", "Print_ISSN","Online_ISSN", "Reporting_Period_Total", "Sourcefilename"]

#dataframe_titles_C_4 = dataframe_titles_C_4.append(dataframe_titles_C_5)

master = master.append(dataframe_titles_C_4)
master = master.append(dataframe_titles_C_5)

master_for_subject = master
print("\n Ungefilterte Mastertabelle: \n",master, "\n")
master.to_csv("master_unfiltered.csv", index=False) # Enthält noch die Medizintitel und irrelevante Zeilen ohne Reporting_Period_Total sind noch enthalten
JSTOR =  master[master.Platform == "JSTOR"] 


# what happens here
# Changes the content of the "Title" column to lowercase
master["Title"] = master["Title"].str.lower()
dataframe_exclude["Title"] = dataframe_exclude["Title"].str.lower()
list = dataframe_exclude["Title"].tolist()
master = master.loc[~master['Title'].isin(list)] # Removes the titles found in files saved in the "exclude" directory



emptyReporting_Period_Total_values = master.loc[master['Reporting_Period_Total'].isna(), :] # List containing the rows with empty Reporting_Period_Total

# Removes rows in which Reporting_Period_Total is empty
#master = master.loc[~master['Reporting_Period_Total'].isna()]
master = remove_empty_rows(master, 'Reporting_Period_Total')

print("Deleted Rows with empty Reporting_Period_Total: \n", emptyReporting_Period_Total_values, "\n")


# removes trailing zero which cause issues with excel 
master["Reporting_Period_Total"] = master["Reporting_Period_Total"].astype(int)         


print("Finale Masterliste: \n", master, "\n")
master.to_csv("master_without_excluded_titles.csv", index=False)







#Parsing List of Single Journals

#os.chdir('./single_journals') # Opens single_journal directory
# for x in listOfSingleJournalFiles: #loops through the list of files in the single_journal directory, and appends each to dataframe_titles_C_4
#     try:
#         i = pd.read_excel(
#             x,
#             usecols = ["Title", "Verlag", "Online_ISSN", "Preis 2021"],
#             )
#         dataframe_single_journals = dataframe_single_journals.append(i, ignore_index=True) # Jeder Schleifendurchlauf erweitert die Tabelle um die neuen Werte
#     except Exception as e: errorMessages.append(("Error in single_journals/%s:" %x, e ))

dataframe_single_journals = parse_xlsx("single_journals", [], 0, ["Title", "Verlag", "Online_ISSN", "Preis 2021"])
#os.chdir('../') # changes working directory to main folder 

#print(dataframe_single_journals)
dataframe_single_journals["Title"] = dataframe_single_journals["Title"].str.lower()
df_single_journals_with_ISSN = pd.DataFrame()
df_single_journals_without_ISSN = pd.DataFrame()
df_single_journals_with_ISSN = dataframe_single_journals.loc[~dataframe_single_journals['Online_ISSN'].isna(), :] # List containing the rows with ISSN
df_single_journals_without_ISSN = dataframe_single_journals.loc[dataframe_single_journals['Online_ISSN'].isna(), :] # List containing the rows without ISSN
df_single_journals_without_ISSN.to_csv("Einzelkaufslisteneinträge ohne Online_ISSN.csv", index=False)


df_print_ISSN = pd.DataFrame()
df_print_ISSN = df_single_journals_with_ISSN.merge(right=master, how="inner", left_on=["Online_ISSN"], right_on=["Print_ISSN", ] )

df_single_journals_with_ISSN = df_single_journals_with_ISSN.merge(right=master, how="inner", left_on=["Online_ISSN"], right_on=["Online_ISSN", ] )

# Removes all single journal entries without a price for 2021 
df_single_journals_with_ISSN = df_single_journals_with_ISSN.loc[~df_single_journals_with_ISSN['Preis 2021'].isna()]

# Division through 0 is possible if Reporting_Period_Total' is 0, but if Preis 2021/Reporting_Period_Total is 0 it will speak for itself
df_single_journals_with_ISSN['Preis 2021/Reporting_Period_Total'] = df_single_journals_with_ISSN['Preis 2021']/df_single_journals_with_ISSN['Reporting_Period_Total']
df_single_journals_with_ISSN['Preis 2021/Reporting_Period_Total'] = df_single_journals_with_ISSN['Preis 2021/Reporting_Period_Total'].round(2)
print("\nMatched Online_ISSN and calculated Preis 2021/Reporting_Period_Total\n", df_single_journals_with_ISSN)




packages = pd.DataFrame()
packages = pd.read_excel("packages/Pakete und Konsortien.xlsx")





df_print_ISSN = df_print_ISSN.loc[~df_print_ISSN['Preis 2021'].isna()]
df_print_ISSN['Preis 2021/Reporting_Period_Total'] = df_print_ISSN['Preis 2021']/df_print_ISSN['Reporting_Period_Total']
df_print_ISSN['Preis 2021/Reporting_Period_Total'] = df_print_ISSN['Preis 2021/Reporting_Period_Total'].round(2)
print("print_issn matching \n",df_print_ISSN)
#df_print_ISSN = df_print_ISSN.append(df_single_journals_with_ISSN)
df_single_journals_with_ISSN = df_single_journals_with_ISSN.append(df_print_ISSN)
#df_single_journals_with_ISSN = df_single_journals_with_ISSN[ "Preis 2021", "Title_y","Publisher", "Platform","Reporting_Period_Total", "Sourcefilename","Preis 2021/Reporting_Period_Total"]
df_single_journals_with_ISSN.to_csv("Single_journal_price.csv", index=False)

#df_print_ISSN.to_csv("single_journal_print_issn.csv", index=False)


#df_single_journals_with_ISSN
packages_exclude = packages.merge(right=df_single_journals_with_ISSN, how="inner", on="Publisher" )
print("\nMaster 1: \n", master)
packages_exludeList = packages_exclude["Online_ISSN"].tolist() #isin()


#match on ISSN, remove all entries with matching ISSN 
#master = master.loc[~master['Online_ISSN'].isin(packages_exludeList)] 
master = master[master.Online_ISSN.isin(packages_exludeList) == False]
print("\nMaster 2: \n", master)


print("\n JSTOR \n",JSTOR) # SUM of
sum_JSTOR_reporting_period_total = pd.DataFrame({"Publisher": ["JSTOR"], "Reporting_Period_Total": [0]})
sum_JSTOR_reporting_period_total["Reporting_Period_Total"] = JSTOR["Reporting_Period_Total"].sum()
print("SUM: \n", sum_JSTOR_reporting_period_total)


# Remove all entries with platform = JSTOR
master = master[master.Platform  != "JSTOR"] 
print("\n Master 3: \n", master)

# Remove all entries where publisher != publisher from packages list
packages_publisher_list = packages["Publisher"].tolist()
print(packages_publisher_list)
master = master[master.Publisher.isin(packages_publisher_list) == True]
print("\nMaster 4: \n", master)
master.to_csv("master_publisher_packages.csv", index=False)

master = master.append(sum_JSTOR_reporting_period_total, ignore_index=True)
# Sum reporting period total for each remaining publisher
master = master.groupby("Publisher")["Reporting_Period_Total"].sum()
print("\nMaster 5: \n", master)

# create df for each publisher calculate
# price packages list / sum RPT
# rounding 2 decimals after 
#take packages as base, add columns RPT for each, and then add calculated result

calculatedPricePackage = packages.merge(right=master, how="left", on="Publisher" )
calculatedPricePackage['Preis 2021/Reporting_Period_Total'] = calculatedPricePackage["Preise 2021"]/calculatedPricePackage['Reporting_Period_Total']
calculatedPricePackage['Preis 2021/Reporting_Period_Total'] = calculatedPricePackage['Preis 2021/Reporting_Period_Total'].round(2)
calculatedPricePackage["Reporting_Period_Total"] = calculatedPricePackage["Reporting_Period_Total"].astype(int)
print("\n calculatedPrice: \n", calculatedPricePackage)

calculatedPricePackage.to_csv("PricePackage.csv", index=False)

#Print Errors, if any
if (len(errorMessages) > 0 and len(errorMessages) < 2):
    print("\033[0;31m", "An Error occured: \n", errorMessages, "\033[0m")
if (len(errorMessages) > 1 ):
    print("\033[0;31m", "Multiple Errors occured: \n", errorMessages, "\033[0m")

# print(dataframe_titles_C_5)
# print("Dataframe_exclude: \n",dataframe_exclude)

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

# print("\n Ungefilterte master_for_subject: \n",master_for_subject, "\n")
# master_for_subject.to_csv("master_for_subject_unfiltered.csv", index=False)
#master_for_subject = remove_empty_rows(master_for_subject, 'Reporting_Period_Total')
# master_for_subject.to_csv("master_for_subject_filtered.csv", index=False)
# print("\n gefilterte master_for_subject: \n",master_for_subject, "\n")

