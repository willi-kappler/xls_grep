# xls_grep
A simple (quick and dirty) tool to search for a given text through a bunch of Excel files

This tool allows you to search for a specific string in a bunch of Excel files.
I wrote this in order to recover important sheets from a crashed hard drive. Use phorotec first then run
xls_grep on the resulting "recup_dir" folders.

## Usage

This tool needs Python 3 (testd with Python 3.5), python-xlrd and python-openpyxl.

`
  python3 xls_grep.py --expression Salary
`

The above command searches recursively for the string "Salary" in the current folder (default) and all its sub folders in all *.xls and *.xlsx files
