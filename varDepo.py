'''
READ ME:

This file contains saved variables/constants for a script called compare_two_xlsx_files.py

FIRST_FILE_ADDRESS and SECOND_FILE_ADDRESS containt adressess of the two xlsx files.

KEY_COLUMN represents the id column of both files, it is asumed that the column is the same in both files, preferably column 'A'.

FIRST_VALUE_COLUMN and SECOND_VALUE_COLUMN contain the letters of the columns holding the values that are to be compared. The order of column letters in both lists should be the same, meaning that the columns with the same data should have the same index.
'''
variables=[{'FIRST_FILE_ADDRESS': '/home/mladen/Documents/PythonVežbe/code_in_place_final/HR _ Vacations.xlsx',
  'FIRST_VALUE_COLUMN': ['G', 'J'],
  'KEY_COLUMN': 'A',
  'SECOND_FILE_ADDRESS': '/home/mladen/Documents/PythonVežbe/code_in_place_final/AllowanceReport_export_2021-06-22.xlsx',
  'SECOND_VALUE_COLUMN': ['J', 'H'],
  'FIRST_LIST_DISPLAY_NAME': 'Google Docs',
  'SECOND_LIST_DISPLAY_NAME': 'Absence'}]
