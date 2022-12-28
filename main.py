'''Module providing basic os functions'''
import os
import pandas as pd

## Constants
YEAR = 2021

## Function Declarations

def create_array_of_xlsx_filenames (directory) :
    '''takes a directoryname as input and returns a list containing 
    the names of all xlsx files in that directory'''
    names = []
    for files in os.listdir(directory):
        if files.endswith('.xlsx'):
            names.append(files)
    return names

def parse_xlsx (directory,  skiprows=None, header= 0,  usecolumns = None, assign = False ):
    '''parses xlsx documents, directory is a mandatory parameter and is defined relative to
    the directory the script is running in, the other ones are optional'''
    dataframe = pd.DataFrame()

    list_of_files = create_array_of_xlsx_filenames(directory)
    print(list_of_files)
    os.chdir(directory)

    for file in list_of_files:
        try:
            i = pd.read_excel(
                file,
                header=header, # Pandas starts the count with 0, unlike Excel which starts at 1
                skiprows=skiprows,
                usecols = usecolumns,
                )
            if assign:
                i=i.assign(Sourcefilename=file)
            dataframe = dataframe.append(i, ignore_index=True)

        except RuntimeError as e:
            errorMessages.append(("Error in ", directory, "/", file, ": ",  e ))
    os.chdir("../")

    return dataframe

def filter_rows (dataframe, column: str, remove_empty: bool):
    '''removes rows in which the specified column is empty or not empty,
    depending on parameter remove_empty'''
    if remove_empty :
        dataframe = dataframe[~dataframe[column].isna()]
    elif not remove_empty :
        dataframe = dataframe[dataframe[column].isna()]

    return dataframe

def sum_jstor_reporting_period_total (dataframe) :
    '''sums up the reporting_period_total values'''
    print("\n jstor_titles \n",dataframe) # SUM of
    return_dataframe = pd.DataFrame({"Publisher": ["JSTOR"], "Reporting_Period_Total": [0]})
    return_dataframe["Reporting_Period_Total"] = dataframe["Reporting_Period_Total"].sum()
    print("SUM: \n", return_dataframe)

    return return_dataframe

def list_of_titles (directory: str) :
    '''returns a list containing the values of the Title column from all xlsx documents in the given directory'''
    dataframe = parse_xlsx(directory)

    return dataframe["Title"].str.lower().tolist()

def calculate_print_single_journal_prices (dataframe) :
    '''takes the filtered masterlist as input, calculates the single journal prices,
    prints them and also returns a dataframe with the calculated prices'''

    dataframe_single_journals = pd.DataFrame()
    dataframe_single_journals = parse_xlsx("single_journals", [], 0, ["Title", "Verlag", "Online_ISSN", f"Preis {YEAR}"])

    #print(dataframe_single_journals)
    dataframe_single_journals["Title"] = dataframe_single_journals["Title"].str.lower()
    df_single_journals_with_issn = df_single_journals_without_issn = df_print_issn = pd.DataFrame()
    df_single_journals_with_issn = dataframe_single_journals[~dataframe_single_journals['Online_ISSN'].isna()] # List containing the rows with ISSN
    df_single_journals_without_issn = dataframe_single_journals[dataframe_single_journals['Online_ISSN'].isna()] # List containing the rows without ISSN
    df_single_journals_without_issn.to_csv("outputs/Einzelkaufslisteneinträge ohne Online_ISSN.csv", index=False)

    df_print_issn = df_single_journals_with_issn.merge(right=dataframe, how="inner", left_on=["Online_ISSN"], right_on=["Print_ISSN", ] )

    df_single_journals_with_issn = df_single_journals_with_issn.merge(right=dataframe, how="inner", left_on=["Online_ISSN"], right_on=["Online_ISSN", ] )

    # Removes all single journal entries without a price for 2021
    df_single_journals_with_issn = df_single_journals_with_issn[~df_single_journals_with_issn[f'Preis {YEAR}'].isna()]

    # Division through 0 is possible if Reporting_Period_Total' is 0, but if Preis {YEAR}/Reporting_Period_Total is 0 it will speak for itself
    df_single_journals_with_issn[f'Preis {YEAR}/Reporting_Period_Total'] = df_single_journals_with_issn[f'Preis {YEAR}']/df_single_journals_with_issn['Reporting_Period_Total']
    df_single_journals_with_issn[f'Preis {YEAR}/Reporting_Period_Total'] = df_single_journals_with_issn[f'Preis {YEAR}/Reporting_Period_Total'].round(2)
    print(f"\nMatched Online_ISSN and calculated Preis {YEAR}/Reporting_Period_Total\n", df_single_journals_with_issn)

    df_print_issn = df_print_issn[~df_print_issn[f'Preis {YEAR}'].isna()]
    df_print_issn[f'Preis {YEAR}/Reporting_Period_Total'] = df_print_issn[f'Preis {YEAR}']/df_print_issn['Reporting_Period_Total']
    df_print_issn[f'Preis {YEAR}/Reporting_Period_Total'] = df_print_issn[f'Preis {YEAR}/Reporting_Period_Total'].round(2)
    print("print_issn matching \n",df_print_issn)
    #df_print_issn = df_print_issn.append(df_single_journals_with_issn)
    df_single_journals_with_issn = df_single_journals_with_issn.append(df_print_issn)
    df_single_journals_with_issn.to_csv("outputs/Single_journal_price.csv", index=False)
    return df_single_journals_with_issn

errorMessages = []
dataframe_titles_C_4 = dataframe_titles_C_5 = dataframe_exclude = master = pd.DataFrame()

column_names_C_4 = ["Title", "Publisher", "Platform", "Print ISSN", "Online ISSN", "Reporting_Period_Total"]
dataframe_titles_C_4 = parse_xlsx("C_4", [8], 7, column_names_C_4, assign=True)

column_names_C_5 = ["Title", "Publisher", "Platform", "Print_ISSN","Online_ISSN", "Metric_Type", "Reporting_Period_Total"]
dataframe_titles_C_5 = parse_xlsx("C_5", [], 14, column_names_C_5, assign=True)

#removes all rows in which Metric_Type isnt Unique_Item_Request
dataframe_titles_C_5 = dataframe_titles_C_5[dataframe_titles_C_5['Metric_Type'] == "Unique_Item_Requests"]

# Columns are named differently in C_4, so I have to rename them to be the same as in the C_5 Standard
dataframe_titles_C_4.columns=["Title", "Publisher", "Platform", "Print_ISSN","Online_ISSN", "Reporting_Period_Total", "Sourcefilename"]

master = master.append(dataframe_titles_C_4)
master = master.append(dataframe_titles_C_5)

# print("\n Ungefilterte Mastertabelle: \n",master, "\n")
master.to_csv("outputs/master_unfiltered.csv", index=False) # Enthält noch die Medizintitel und irrelevante Zeilen ohne Reporting_Period_Total sind noch enthalten

## Calculating

sum_JSTOR = sum_jstor_reporting_period_total(master[master.Platform == "JSTOR"])


# Changes the content of the "Title" column to lowercase
master["Title"] = master["Title"].str.lower()

# Removes the titles found in files saved in the "exclude" directory
master_filtered = master[~master['Title'].isin(list_of_titles("exclude"))]

emptyReporting_Period_Total_values = filter_rows(master, 'Reporting_Period_Total', False)
print("Deleted Rows with empty Reporting_Period_Total: \n", emptyReporting_Period_Total_values, "\n")

# Removes rows in which Reporting_Period_Total is empty
master_filtered = filter_rows(master_filtered, 'Reporting_Period_Total', True)

# removes trailing zero which cause issues with excel
master_filtered["Reporting_Period_Total"] = master_filtered["Reporting_Period_Total"].astype(int)


# print("Finale Masterliste: \n", master_filtered, "\n")
master_filtered.to_csv("outputs/master_without_excluded_titles.csv", index=False)


single_journals_with_ISSN = calculate_print_single_journal_prices(master_filtered)
#df_print_issn.to_csv("single_journal_print_issn.csv", index=False)

packages = pd.DataFrame()
packages = pd.read_excel("packages/price_packages/Pakete und Konsortien.xlsx")
#df_single_journals_with_issn
packages_exclude = packages.merge(right=single_journals_with_ISSN, how="inner", on="Publisher" )
print("\nMaster 1: \n", master_filtered)
packages_exludeList = packages_exclude["Online_ISSN"].tolist() #isin()


#match on ISSN, remove all entries with matching ISSN
master_filtered = master_filtered[~master_filtered.Online_ISSN.isin(packages_exludeList)]
print("\nMaster 2: \n", master_filtered)

# Remove all entries with platform = JSTOR
master_filtered = master_filtered[master_filtered.Platform  != "JSTOR"]
print("\n Master 3: \n", master_filtered)

# Remove all entries where publisher != publisher from packages list
packages_publisher_list = packages["Publisher"].tolist()
print(packages_publisher_list)
master_filtered = master_filtered[master_filtered.Publisher.isin(packages_publisher_list)]
print("\nMaster 4: \n", master_filtered)
master_filtered.to_csv("outputs/master_publisher_packages.csv", index=False)

master_filtered = master_filtered.append(sum_JSTOR, ignore_index=True)
# Sum reporting period total for each remaining publisher
master_filtered = master_filtered.groupby("Publisher")["Reporting_Period_Total"].sum()
print("\nMaster 5: \n", master_filtered)

# create df for each publisher calculate
# price packages list / sum RPT
# rounding 2 decimals after
#take packages as base, add columns RPT for each, and then add calculated result

calculatedPricePackage = packages.merge(right=master_filtered, how="left", on="Publisher" )
calculatedPricePackage[f'Preis {YEAR}/Reporting_Period_Total'] = calculatedPricePackage["Preise 2021"]/calculatedPricePackage['Reporting_Period_Total']
calculatedPricePackage[f'Preis {YEAR}/Reporting_Period_Total'] = calculatedPricePackage[f'Preis {YEAR}/Reporting_Period_Total'].round(2)
calculatedPricePackage["Reporting_Period_Total"] = calculatedPricePackage["Reporting_Period_Total"].astype(int)
print("\n calculatedPrice: \n", calculatedPricePackage)

calculatedPricePackage.to_csv("outputs/PricePackage.csv", index=False)

#Print Errors, if any
if (len(errorMessages) > 0 and len(errorMessages) < 2):
    print("\033[0;31m", "An Error occured: \n", errorMessages, "\033[0m")
if len(errorMessages) > 1 :
    print("\033[0;31m", "Multiple Errors occured: \n", errorMessages, "\033[0m")
