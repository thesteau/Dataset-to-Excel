# Dataset-to-Excel
Aggregates multiple folders of CSV or TSV files into an Excel file for data analysis.

## Prerequisites
```
Python 3.7 or later 3.x
pandas 1.3.2
xlxswriter (latest stable release)
```

## Getting Started

After you have the required packages installed, please update the "partner_dictionary" to include your subdirectories under the ```the_partners()``` method.
```python
partner_dictionary = {
               "SampleData": 0,
               "SkippingSample", 2
         }
```
Note: Each key and value corresponds to the following:  
```
{STRING - parner name (Must match the folder name): INTEGER - skip heading rows (Enter 0, otherwise any integer greater than 0)}
```
Please ensure that this is the following folder structure:
```graphql
C:\SampleFolder
  ├─ SampleData
  │  ├─ randompart1.csv
  │  └─ randompart2.csv   
  └─ SkippingSample
     ├─ skipSample.tsv    # Please have one filetype per folder if the data's structure differs 
     └─ skipSample2.csv   # (Such as headings, columns, etc.)
```
---
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
---
Either case, you are going to be prompted with the following:
```
Enter folder source path: 
```
Enter the appropriate full folder path with subdirectories of files (CSV, TSV):
```Mac
Enter folder source path: /SampleFolder
```
Or for Windows
```Windows
Enter folder source path: C:\SampleFolder
```

## Results
On the entered path, you will find an Excel file with the following merged data with tabs seperated by the subdirectories:

![image](https://raw.githubusercontent.com/thesteau/Dataset-to-Excel/main/images/Capture.PNG)

## Motivation
This program was created when I was an analyst. We had a ton of files from several partners (Such as Apple, Microsoft, etc). Humanly, processing all the data would take hours so I had developed this program to automate the aggregation process with the pandas library to produce an Excel file.

## Author
Steven Au

## License
See the LICENSE.md for details
