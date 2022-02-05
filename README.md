# CSV_XLSX_Merge

## Prerequisites
requires installation  pandas, openpyxl, IPython (for instance via 
```
pip install pandas openpyxl IPython
```

## How to Use:

Save the counter 5 release source xlsx in the xlsx directory.  

Save the xlsx containing the titles that should be excluded in the exclude directory, ensure that the first column is named Titles and contains all the titles.

## Results

Currently 2 csv-Files will be generated, one containing all the titles from the source files, and one where the titles from exclude are filtered out.