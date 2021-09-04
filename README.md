# Dataset-to-Excel
Aggregvates multiple folders of CSV or TSV files into an Excel file for data analysis.

# TO DO
Add screenshots and document how to use the code.  

## Prerequisites
```
Python 3.7 or later
Pandas 1.2.3
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
Or for widnows
```Windows
Enter folder source path: C:\SampleFolder
```
## Usage



## Author
Steven Au

## License
See the LICENSE.md for details

## Acknowledgements
Special thanks to the original developer of the Pantab module!
