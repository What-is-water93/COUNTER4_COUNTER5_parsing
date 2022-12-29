"""Module providing basic os functions"""
import os
import pandas as pd

# Constants
YEAR = 2021
COLUMN_NAMES_C_4 = ["Title", "Publisher", "Platform", "Print ISSN", "Online ISSN", "Reporting_Period_Total"]
COLUMN_NAMES_C_5 = ["Title", "Publisher", "Platform", "Print_ISSN", "Online_ISSN", "Metric_Type", "Reporting_Period_Total"]
COLUMNS_PACKAGE_FILES = ["Title_Paketliste", "Title_Master", "Online_ISSN", "Print_ISSN", "Publisher", "Platform", "Reporting_Period_Total", "Sourcefilename"]
ERROR_COLOUR = "\033[0;31m"
DEFAULT_COLOUR = "\033[0m"

# Function Declarations


def create_array_of_xlsx_filenames(directory: str):
    """takes a directoryname as input and returns a list containing
    the names of all xlsx files in that directory"""
    names = []
    for files in os.listdir(directory):
        if files.endswith(".xlsx"):
            names.append(files)
    return names


def parse_xlsx(directory: str, skiprows: int = None, header: int = 0, usecolumns=None, assign: bool = False):
    """parses xlsx documents, directory is a mandatory parameter and is defined
    relative to the directory the script is running in, the other ones are
    optional. Rows are 0 indexed, ie. excel row number 8 is 7 for this tool"""
    dataframe = pd.DataFrame()

    list_of_files = create_array_of_xlsx_filenames(directory)
    print(list_of_files)
    os.chdir(directory)

    for file in list_of_files:
        try:
            i = pd.read_excel(
                file,
                header=header,  # Pandas starts the count with 0, unlike Excel which starts at 1
                skiprows=skiprows,
                usecols=usecolumns,
            )
            if assign:
                i = i.assign(Sourcefilename=file)
            dataframe = dataframe.append(i, ignore_index=True)

        except Exception as e:  # pylint: disable=broad-except
            error_messages.append((f"Error in {directory}/{file}: {e}"))
    os.chdir("../")

    return dataframe


def remove_empty_rows(dataframe, column):
    """removes rows where the given column is empty"""
    return dataframe[~dataframe[column].isna()]


def remove_not_empty_rows(dataframe, column):
    """removes rows where the given column is not empty"""
    return dataframe[dataframe[column].isna()]


def sum_jstor_reporting_period_total(dataframe):
    """sums up the reporting_period_total values"""
    print("\n jstor_titles \n", dataframe)  # SUM of
    return_dataframe = pd.DataFrame({"Publisher": ["JSTOR"], "Reporting_Period_Total": [0]})
    return_dataframe["Reporting_Period_Total"] = dataframe["Reporting_Period_Total"].sum()
    print("SUM: \n", return_dataframe)

    return return_dataframe


def list_of_titles(directory: str):
    """returns a list containing the values of the Title column from all xlsx documents in the given directory"""
    dataframe = parse_xlsx(directory)

    return dataframe["Title"].str.lower().tolist()


def calculate_print_single_journal_prices(dataframe):
    """takes the filtered masterlist as input, calculates the single journal prices,
    prints them and also returns a dataframe with the calculated prices"""

    dataframe_single_journals = parse_xlsx("single_journals", [], 0, ["Title", "Verlag", "Online_ISSN", f"Preis {YEAR}"])

    dataframe_single_journals["Title"] = dataframe_single_journals["Title"].str.lower()
    df_single_journals_with_issn = remove_empty_rows(dataframe_single_journals, "Online_ISSN")  # List containing the rows with ISSN
    df_single_journals_without_issn = remove_not_empty_rows(dataframe_single_journals, "Online_ISSN")  # List containing the rows without ISSN
    df_single_journals_without_issn.to_excel("outputs/Einzelkaufslisteneinträge ohne Online_ISSN.xlsx", index=False)

    df_print_issn = \
        df_single_journals_with_issn.merge(right=dataframe, how="inner", left_on=["Online_ISSN"], right_on=["Print_ISSN", ])

    df_single_journals_with_issn = \
        df_single_journals_with_issn.merge(right=dataframe, how="inner", left_on=["Online_ISSN"], right_on=["Online_ISSN", ])

    # Removes all single journal entries without a price for 2021
    df_single_journals_with_issn = remove_empty_rows(df_single_journals_with_issn, f"Preis {YEAR}")

    # Division through 0 is possible if Reporting_Period_Total" is 0, but if Preis {YEAR}/Reporting_Period_Total is 0 it will speak for itself
    df_single_journals_with_issn[f"Preis {YEAR}/Reporting_Period_Total"] = \
        (df_single_journals_with_issn[f"Preis {YEAR}"]/df_single_journals_with_issn["Reporting_Period_Total"]).round(2)

    df_print_issn = remove_empty_rows(df_print_issn, f"Preis {YEAR}")

    df_print_issn[f"Preis {YEAR}/Reporting_Period_Total"] = \
        (df_print_issn[f"Preis {YEAR}"]/df_print_issn["Reporting_Period_Total"]).round(2)
    # print("print_issn matching \n", df_print_issn)
    df_single_journals_with_issn = df_single_journals_with_issn.append(df_print_issn)
    df_single_journals_with_issn.to_excel("outputs/Single_journal_price.xlsx", index=False)

    return df_single_journals_with_issn


def print_packages_calculate_price(dataframe):
    """merges each package list with the masterlist and prints the resulting table.
    Sum is not calculated because the merge often contain titles from other packages which requires a manual check"""
    list_of_package_files = create_array_of_xlsx_filenames("packages")
    for title in list_of_package_files:
        try:
            package = pd.read_excel(f"packages/{title}", usecols=["Title", "Online_ISSN"])
            package = package.drop_duplicates(subset=["Online_ISSN"], ignore_index=True)
            merged = package.merge(dataframe.dropna(subset=["Online_ISSN"]), on=["Online_ISSN"], how="inner", suffixes=("_Paketliste", "_Master"))
            merged.to_excel(f"outputs/packages/{title}", columns=COLUMNS_PACKAGE_FILES, index=False)
        except Exception as e:  # pylint: disable=broad-except
            error_messages.append((f"Error in packages/{title}: {e}"))


error_messages = []  # Contains errors that occured while parsing input .xlsx files with error message, directory and file

dataframe_titles_C_4 = parse_xlsx("C_4", [8], 7, COLUMN_NAMES_C_4, assign=True)
dataframe_titles_C_5 = parse_xlsx("C_5", [], 14, COLUMN_NAMES_C_5, assign=True)

# removes all rows in which Metric_Type isnt Unique_Item_Request
dataframe_titles_C_5 = dataframe_titles_C_5[dataframe_titles_C_5["Metric_Type"] == "Unique_Item_Requests"]

# Columns are named differently in C_4, so I have to rename them to be the same as in the C_5 Standard
dataframe_titles_C_4.columns = ["Title", "Publisher", "Platform", "Print_ISSN", "Online_ISSN", "Reporting_Period_Total", "Sourcefilename"]

master = pd.DataFrame()
master = master.append(dataframe_titles_C_4)
master = master.append(dataframe_titles_C_5)

master.to_excel("outputs/master_unfiltered.xlsx", index=False)  # Enthält noch die Medizintitel und irrelevante Zeilen ohne Reporting_Period_Total sind noch enthalten

# Calculating
sum_JSTOR = sum_jstor_reporting_period_total(master[master.Platform == "JSTOR"])

# Changes the content of the "Title" column to lowercase
master["Title"] = master["Title"].str.lower()

# Removes the titles found in files saved in the "exclude" directory
master_filtered = master[~master["Title"].isin(list_of_titles("exclude"))]

emptyReporting_Period_Total_values = remove_not_empty_rows(master, "Reporting_Period_Total")
print("Deleted Rows with empty Reporting_Period_Total: \n", emptyReporting_Period_Total_values, "\n")

# Removes rows in which Reporting_Period_Total is empty
master_filtered = remove_empty_rows(master_filtered, "Reporting_Period_Total")

# removes trailing zero which cause issues with excel
master_filtered["Reporting_Period_Total"] = master_filtered["Reporting_Period_Total"].astype(int)

# print("Finale Masterliste: \n", master_filtered, "\n")
master_filtered.to_excel("outputs/master_without_excluded_titles.xlsx", index=False)

single_journals_with_ISSN = calculate_print_single_journal_prices(master_filtered)
# df_print_issn.to_csv("single_journal_print_issn.csv", index=False)

# packages = pd.DataFrame()
# packages = pd.read_excel("packages/price_packages/Pakete und Konsortien.xlsx")

# Removes all entries from JSTOR from the masterlist
master_filtered_no_jstor = master_filtered[master_filtered.Platform != "JSTOR"]

print_packages_calculate_price(master_filtered_no_jstor)

# Print Errors, if any
if len(error_messages) == 1:
    print(ERROR_COLOUR, "An Error occured: \n", error_messages, DEFAULT_COLOUR)  # colours as variables?
if len(error_messages) > 1:
    print(ERROR_COLOUR, "Multiple Errors occured: \n", error_messages, DEFAULT_COLOUR)
