"""Module providing basic os functions"""
import os
import pandas as pd

# Constants
YEAR = 2022
COLUMN_NAMES_C_4 = ["Title", "Publisher", "Platform", "Print ISSN", "Online ISSN", "Reporting_Period_Total"]
COLUMN_NAMES_C_5 = ["Title", "Publisher", "Platform", "Print_ISSN", "Online_ISSN", "Metric_Type", "Reporting_Period_Total"]
COLUMNS_PACKAGE_FILES = ["Title_Paketliste", "Title_Master", "Online_ISSN", "Print_ISSN", "Publisher", "Platform", "Reporting_Period_Total", "Sourcefilename"]
RED = "\033[0;31m"
GREEN = '\033[92m'
BOLD = '\033[1m'
RESET_TERMINAL_STYLE = "\033[0m"

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


def remove_empty_rows(dataframe, column: str):
    """removes rows where the given column is empty"""

    return dataframe[~dataframe[column].isna()]


def remove_not_empty_rows(dataframe, column: str):
    """removes rows where the given column is not empty"""

    return dataframe[dataframe[column].isna()]


def print_sum_of_reporting_period_total_jstor(dataframe):
    """sums up the reporting_period_total values"""

    print(BOLD, "The Reporting_Period_Total sum for JSTOR is:", GREEN, dataframe["Reporting_Period_Total"].sum().astype(int), RESET_TERMINAL_STYLE)


def calculate_price_print_single_journal(dataframe):
    """takes the filtered masterlist as input, calculates the single journal prices,
    prints them and also returns a dataframe with the calculated prices"""

    dataframe_single_journals = parse_xlsx("single_journals")
    dataframe_single_journals["Title"] = dataframe_single_journals["Title"].str.lower()
    df_single_journals_with_issn = remove_empty_rows(dataframe_single_journals, "Online_ISSN")  # List containing the rows with ISSN
    df_single_journals_without_issn = remove_not_empty_rows(dataframe_single_journals, "Online_ISSN")  # List containing the rows without ISSN
    df_single_journals_without_issn.to_excel("outputs/Einzelkaufslisteneinträge ohne Online_ISSN.xlsx", index=False)

    df_print_issn = \
        df_single_journals_with_issn.merge(right=dataframe, how="inner", left_on=["Online_ISSN"], right_on=["Print_ISSN", ])

    df_single_journals_with_issn = \
        df_single_journals_with_issn.merge(right=dataframe, how="inner", on=["Online_ISSN"])

    # Removes all single journal entries without a price for 2021
    df_single_journals_with_issn = remove_empty_rows(df_single_journals_with_issn, f"Preis {YEAR}")

    # If the new column has the value "inf" a division through 0 took place (happens if Reporting_Period_Total = 0)
    df_single_journals_with_issn[f"Preis {YEAR}/Reporting_Period_Total"] = \
        (df_single_journals_with_issn[f"Preis {YEAR}"]/df_single_journals_with_issn["Reporting_Period_Total"]).round(2)

    df_print_issn = remove_empty_rows(df_print_issn, f"Preis {YEAR}")

    df_print_issn[f"Preis {YEAR}/Reporting_Period_Total"] = \
        (df_print_issn[f"Preis {YEAR}"]/df_print_issn["Reporting_Period_Total"]).round(2)

    df_single_journals_with_issn = df_single_journals_with_issn.append(df_print_issn)
    df_single_journals_with_issn.drop(columns=["Title_y", "Metric_Type"]).to_excel("outputs/Single_journal_price.xlsx", index=False)


def calculate_price_print_kombi(dataframe):
    """calculates the price/Reporting_Period_Total with the Reporting_Period_Total
    from masterlist and the prices from the file in the kombi_abo_price directory"""

    kombi = parse_xlsx("kombi_abo_preise")
    kombi = remove_empty_rows(kombi, "ISSN")
    kombi = kombi.merge(right=dataframe, how="inner", left_on=["ISSN"], right_on=["Online_ISSN", ])
    kombi["Preis/Reporting_Period_Total"] = (kombi["Preis"]/kombi["Reporting_Period_Total"]).round(2)

    column_order = ['Titel', 'ISSN', 'Verlag', 'Preis', 'Reporting_Period_Total', 'Preis/Reporting_Period_Total',
                    'Bestellzeichen', 'sonst. Bemerkung', 'Publisher', 'Platform', 'Print_ISSN', 'Online_ISSN',  'Sourcefilename']
    kombi.drop(columns=["Title", "Metric_Type"]).loc[:, column_order].to_excel("outputs/kombi.xlsx", index=False)


def print_packages_as_xlsx(dataframe):
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

master.to_excel("outputs/master_with_empty_RPT.xlsx", index=False)  # Enthält irrelevante Zeilen ohne Reporting_Period_Total

# Calculating


# Changes the content of the "Title" column to lowercase
master["Title"] = master["Title"].str.lower()

emptyReporting_Period_Total_values = remove_not_empty_rows(master, "Reporting_Period_Total")
print("Deleted Rows with empty Reporting_Period_Total: \n", emptyReporting_Period_Total_values, "\n")

# Removes rows in which Reporting_Period_Total is empty
master_filtered = remove_empty_rows(master, "Reporting_Period_Total")

print_sum_of_reporting_period_total_jstor(master_filtered[master_filtered.Platform == "JSTOR"])

master_filtered.to_excel("outputs/master_without_empty_RPT.xlsx", index=False)

calculate_price_print_single_journal(master_filtered)

# Kombiabopreise
calculate_price_print_kombi(master_filtered)

# Removes all entries from JSTOR from the masterlist
master_filtered_no_jstor = master_filtered[master_filtered.Platform != "JSTOR"]

print_packages_as_xlsx(master_filtered_no_jstor)

# Print Errors, if any
if len(error_messages) == 1:
    print(RED, "An Error occured: \n", error_messages, RESET_TERMINAL_STYLE)  # colours as variables?
if len(error_messages) > 1:
    print(RED, "Multiple Errors occured: \n", error_messages, RESET_TERMINAL_STYLE)
