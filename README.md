# COUNTER4_COUNTER5_parsing

This script aims to simplify statistical analysis of counter 4 and counter 5 reports.  
It creates a main table file with all entries from n amount of reports. 
Additionally it calculates prices per request for journals that are singularly bought or are part of a combo abo.  
Lastly it creates files containing the request numbers for all media that are part of a package deal, based on title lists for those packages

## Prerequisites
The script requires some dependencies, they can either be installed via the dependency manager [poetry](https://github.com/python-poetry/poetry), which has to be run in the script directory:
```
poetry install
```
or via pip, which can be run anywhere as long it is installed
```
pip install pandas openpyxl
```

## How to Use:
### Preparations
1. Save the counter 4 source xlsx files in the C_4 directory. 
2. Save the counter 5 release source xlsx files in the C_5 directory. 
3. Save a list with titles that are bought singularly in the single_journals directory
4. Save a list with titles that are bought as part of a combo-abo in the combo_abo_prices directory
5. Save the title list for each package in the package directory


### Run the script:
Open a terminal and navigate to the directory of the script:
```
cd PATH_TO_DIRECTORY
```
Run the script, depending on your OS/Pythonversion it's for instance
```
python3 main.py
```
or
```
python main.py
```


## Results

The script saves all files it generates in the outputs directory.
### The following files are always created:
- **main.xlsx**
	- main table, contains all the parsed C_4 and C_5 report entries, on condition that they have a value for the  `reporting_period_total` column (this filters out rows where publishers added a comment in the first column but left everything else empty)
- **combo.xlsx**
	- this contains all entries with an ISSN number from files in the `combo_abo_prices` directory. The script joined the entries via ISSN  on Online_ISSN with the publisher data, adding the extra columns `reporting_period_total` and `Preis/Reporting_Period_Total` filled with the correct data.
- **combo_print.xlsx**
	- this contains all entries with an ISSN number from files in the `combo_abo_prices` directory. The script joined the entries via ISSN  on Print_ISSN with the publisher data, adding the extra columns `reporting_period_total` and `Preis/Reporting_Period_Total` filled with the correct data.
    Entries from 2 source files are removed, since those publisher used the same ISSN as Online and Print ISSN, leading to double matched data.
- **Single_journal_price.xlsx**
	- similar to combo.xslx,  the script joins on the ISSN (print and online) and adds the  `reporting_period_total` and `Preis/Reporting_Period_Total` columns.

### The following files are only conditionally created:
- **single_journals_without_Online_ISSN.xlsx**
	- If any of the files in the `single_journals` directory contains rows with empty Online_ISSN columns this file will be created with all affected rows
- **main_with_empty_RPT.xlsx**
	- If the main table contains any rows with  an empty ( = no value) `reporting_period_total` column it creates a file containing those rows
