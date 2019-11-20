# Automatically Compare Two Excel or CSV Files 

November 20 2019

10:06 AM

## Overview 
This tool compares excel documents or CSV files column by column and identifies differences.  

This was designed for testing the output of a report in two different environments, but can be used for any excel document compatible with openpyxl*.  If the input files are CSV, they are converted to XLSX before the comparison begins.

Comparisons are done by identifying if values being compared are numbers, dates, or strings.  If numbers, the difference between the values are stored in a difference column.  If dates or strings, the word "Same" or "Different" is stored.  The output file contains each column of the input files (named "left" and "right") alongside the difference column.  The script proceeds until it has compared every sheet in the file.  

The script can use the values of a single column to line up the rows of the input files, so that missing values do not offset the comparison.  

*Note: as of writing there are some limitations of openpyxl when working with Excel documents containing images.

### Key benefits include

1.  Output file contains values being compared side by side with differences, reducing context switching. This means better ease of testing and reduces human error.

1.  Line up values so missing values are identified quickly and do not interfere with the comparison

1.  Compare any excel files quickly and automatically, meaning less developer time wasted doing excel comparisons


## Setup

Prepare Python Environment
--------------------------

The necessary packages are found in the requirements.txt file. Install using pip and you should be good to go.

For example 
`pip install -r requirements.txt `

### Make executable
---------------
Compiling the script to an executable is optional, but can help people who are not familiar with python use the command line tool.  

To deploy as an executable, I recommend using [pyinstaller](https://www.pyinstaller.org/).

Install pyinstaller, then compile the project by executing the following command:

pyinstaller compare.py -F -n compare_excel -i ./icon/icon.ico

1. The -F flag is for a one-file bundled executable
2. The -n flag gives the bundled app a name
3. The -i flag assigns an icon to the application

This will create a directory called "dist" in the project directory with containing your executable. Distribute the executable to your users.

On Windows, you may need to install some microsoft packages, like the VC++ redistributable package, prior to being able to compile the application.

### Using the tool
--------------
For more information about parameters and options, pass the argument "--help" to the tool.

#### Positional arguments
  left                  
  Path to first file for comparison. Can be CSV or XLSX.
                        In output file, these values will be on left
  
  right                 
  Path to second file for comparison. Can be CSV or
                        XLSX. In output file, these values will be on right
  
  output                
  Path to output file. If file exists it will be
                        overwritten. If compare_type is 'sorted' then it will
                        contain copies of data from original files as well as
                        the values side by side in a combined sheet.

optional arguments:
  -h, --help            show this help message and exit
  --threshold THRESHOLD, -t THRESHOLD
                        threshold for numeric values to be considered
                        different. e.g. when threshold = 0.01 if left and
                        right values are closer than 0,01 then consider the
                        same. Mainly affects coloring of difference column for
                        numeric values
  --open OPEN, -p OPEN  if true, open output file on completion using
                        os.system. Output file path must resolve to a file.
                        Adds quotes around file name so that paths with spaces
                        can resolveon windows machines.
  --compare_type COMPARE_TYPE, -c COMPARE_TYPE
                        if set to 'sorted', the comparison tool will attempt
                        to line up each side based on the values of
                        sort_column specified. 'default' is a cell-by-cell
                        comparison.
  --sort_column SORT_COLUMN, -s SORT_COLUMN
                        numeric offset (1-based) of column to use for sorting.
                        E.g. a primary key. if compare type is 'sorted', this
                        column will be used to sort and line up each side
  --has_header HAS_HEADER, -d HAS_HEADER
                        if sheets have headers, set to True so the headers can
                        be excluded from comparison
  --sheet_matching SHEET_MATCHING, -m SHEET_MATCHING
                        can be either 'name' or 'order'. If name, only sheets
                        with the same name are compared. if order, sheets are
                        compared in order. E.g. 1st sheet vs 1st sheet.
  --convert_csv CONVERT_CSV, -v CONVERT_CSV
                        if True, convert csv files to xlsx



1.  Call the utility by passing in 3 file paths for left input, right input, and output files.
1.  if you want input files to be sorted and lined up by a specified column first, you must pass set --compare_type to "sorted" and --sort_column to a numeric value corresponding to the column you want to sort by (1-based, so first column = 1)
2.  If you want to open the output file automatically, set --open to True
3.  If your file does not have headers, pass the arguments --has_headers False


## Troubleshooting

If you get an error like this then you need to install the Microsoft Visual C++ package.

![System Error
Q The program can’t start because api-ms-win-crt-stdio-l1-1-OdIl is
‘ missing from your computer. Try reinstalling the program to fix this
problem.](https://i.imgur.com/eTgqVN4.png)


See [this link](http://www.thewindowsclub.com/api-ms-win-crt-runtime-l1-1-0-dll-is-missing) for more information.



#### Download links

Depending on your machine, you may need one or the other

[32 bit download](http://www.microsoft.com/en-gb/download/details.aspx?id=5555)

[64 bit download](http://www.microsoft.com/en-us/download/details.aspx?id=14632)


