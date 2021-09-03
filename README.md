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

After you have the required packages installed, simply run the script on the same directory as the .hyper file.

However, you may run this script for the specific method as shown below.

```python
import HyperConvert as hc

hyper = tableau_data_path.hyper  # The exact path of the original file with the .hyper extension
dest = destination_path.csv      # The destination path for the file with the .csv extension

hc.hyper_to_csv(hyper, dest)
```

## Usage



## Author
Steven Au

## License
See the LICENSE.md for details

## Acknowledgements
Special thanks to the original developer of the Pantab module!
