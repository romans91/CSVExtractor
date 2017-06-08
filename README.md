# CSV Extractor

Some applications require static table-style data but don't need an entire database connection devoted to this task. Hard coding data into the program is also not an ideal alternative. One solution for this problem is comma separated text or numeric data stored in .csv files, to be read by the program. Microsoft Excel offers ways of creating and managing table-style data but converting that data into .csv files is not always simple. This java executable works with one click, command or .bat file, traverses excel spreadsheets, finds gropus of cells marked as .csv data and creates as many such files as needed. 

# Usage
### Marking CSVs in the Excel spreadsheet
At any row in any sheet, in column A, the start of a .csv is marked by putting the desired output filename into the cell. For example:
```
MyData.csv
```
On column B of the same row, the cell next to the filename has to contain a numeric value specifying the amount of columns in the .csv. For example:
```
3
```
The extractor will take data only from that amount of columns from the .xlsx into the output .csv, starting from column A. This example will produce a .csv with 3 columns, taken only from columns A, B and C of the .xlsx. 
- Any graph, formula or cell that is outside of that range will be ignored by the extractor
- Blank cells inside that range will create a jagged .csv file
- The extractor is not type sensitive
- If a numeric value is divisible by 1, a decimal point will not show up in the output, even if the value is something like "1.0"

The end of the .csv file is marked by a cell in column A that only contains a full stop character. Leaving this out or in a different column to A will not produce a .csv.
```
.
```

Example of cells marked to be a .csv:

|  | A | B | C | D |
| :---: | :-: | :-: | :-: | :-: |
| 1 |   |   |   |   |
| 2 | MyData.csv | 3 |   |   |
| 3 | 1 | 2 | Text 1 |   |
| 4 | 3 | 4 | Text 2  |   |
| 5 | . |   |   |   |
| 6 |   |   |   |   |

This produces a file named "MyData.csv" in a folder named "csv" with the following contents:
```
1,2,Text 1
3,4,Text 2
```

### Running the application
If the jar is simply run, it will look for a file called "Data.xlsx" in its directory and place all the output .csv files in a folder in the same directory called "csv". The result of the above example will be in "csv/MyData.csv", relative to the .jar.

If the compiled .jar file is run with the following command:
```
java -jar CSVExtractor.jar Spreadsheet1.xlsx Spreadsheet2.xlsx
```
Instead of looking for "Data.xlsx", it will go through all the spreadsheets named in the command as arguments, in  this case being "Spreadsheet1.xlsx" and "Spreadsheet2.xlsx". 