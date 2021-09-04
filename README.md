# Dataset-to-Excel
Aggregvates multiple folders of CSV or TSV files into an Excel file for data analysis.

## Prerequisites
```
Python 3.7 or later 3.x
Pandas 1.2.3
xlxswriter (latest stable release)
```

## Getting Started

After you have the required packages installed, please update the "partner_dictionary" to include your subdirectories.
```python
partner_dictionary = {
               "SampleData": 0,
               "SkippingSample", 2
         }
```
Each key and value corresponds to the following:  
```
{STRING - parner name (Must match the folder name): INTEGER - skip heading rows (Enter 0, otherwise any integer greater than 0)}
```
Please ensure that this is the following folder structure:
```
Example: Parent directory is, for Windows, where the parent directory: +, subdirectory: *, and file: ^.
    +C:\downloads
    |          |
    *folder1   *partnerName
    |          |
    ^file.csv  | ^ partnerfile.tsv  # Please have only one filetype per folder if the data's structure differs (Such as headings, columns, etc.)
               | ^ Partner2.csv
```

Afterwards, simply run or import the module. (These two scenarios assume that you are importing this module)
```python
import datatoexcel as dte

dte.DataToExcel().process_data()
```  
or  
```python
import datatoexcel as dte

aggregate = dte.DataToExcel()
aggregate.process_data()
```
Either case, you are going to be prompted with the following:
```
Enter folder source path: 
```
Enter the apprpriate full folder path with subdirectories of files (CSV, TSV):
```Mac
Enter folder source path: /SampleFolder
```
Or for Windows
```Windows
Enter folder source path: C:\SampleFolder
```

## Results
On the entered path, you will find an Excel file with the following merged data with tabs seperated by the subdirectories:
```
![image](https://raw.githubusercontent.com/thesteau/Dataset-to-Excel/main/images/Capture.PNG)
```

## Author
Steven Au

## License
See the LICENSE.md for details
