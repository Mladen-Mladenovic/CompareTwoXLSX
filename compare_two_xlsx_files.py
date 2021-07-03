#! python3
'''
	Takes two .xlsx docs and 
	compares the values for chosen columns.
	The program makes changes directly in the documents.
	THE TWO DOCS ARE DELETED AFTER THE ANALISYS.

	The keys in both documents have to correspond, 
	meaning the id has to match for the rows with the same data.

	The constants here (adresses and column letters) are saved in a file varDepo.py
'''


import send2trash, openpyxl, pprint ,varDepo, logging
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logging.disable(level=logging.CRITICAL)

# Loads xlsx document, requires doc address and sheet name, returnes workbook and worksheet as a tuple.
def load_data(address, sheet):
	wb = openpyxl.load_workbook(address, data_only=True)
	if sheet == None:
		ws = wb.active
	else:
		ws = wb[sheet]
	logging.debug(f'ws= {ws}')
	return wb, ws

# Counts number of cells until the first blank cell, returns the integer
def maximum_row(ws, key_column):
	if type(key_column) == int:
		key_column = get_column_letter(key_column)
	count=0
	for cell in ws [key_column]:
		if cell.value == None: # Stop with the first blank cell
			break
		else:
			count +=1
	return count

# Removes diacritics, strips white space and removes interpunction, changes are made directly in the document.
def format_names(workSheet, col_num):
	ws=workSheet
	maxim = maximum_row (ws, col_num)
	logging.debug(f'maxim ={maxim}')

	for col in ws.iter_cols(min_row=2, min_col=col_num, max_row=maxim, max_col=col_num):
		for cell in col:
			if cell.value == None: # Stop with the first blank cell
				break
			else:
				cell.value = cell.value.strip()
				cell.value = cell.value.translate(str.maketrans({'Č':'C','č':'c','Ć':'C','ć':'c','Š':'S','š':'s','Đ':'Dj','đ':'dj','Ž':'Z','ž':'z','.':''}))

# Creates a list of two dictionaries, one for each dataset, returns a list.
def create_dict_list(ws, key_column ,value_column):
	first = {}
	second = {}
	col_list = [first, second]
	for i in range (len(value_column)):
		count=1
		for cell in ws[key_column]:
			if cell.value == None: # Stop with the first blank cell
				break
			else:
				index = value_column[i]+str(count) # Creates cell name by combining letter from the value column with the number from the >> i << counter
				col_list[i][cell.value] = ws[index].value # Adds value to the dictionary from the cell on the position saved in >> index << 
				count+=1
		col_list.append # Appends the dictionary for a single dataset to the output-ready list
	logging.debug(f'col_list = \n{col_list}')
	return col_list

# Combines data from first name and last name columns, and fills column 'A' with the full name data.
def construct_full_name(ws):
	maxim= ws.max_row
	firstNames = []
	lastNames = []
	# Extract first name
	for col in ws.iter_cols(min_row=2, min_col=1, max_row=maxim, max_col=1):
		for cell in col:
			firstNames.append(cell.value)
	# Extract last name
	for col in ws.iter_cols(min_row=2, min_col=2, max_row=maxim, max_col=2):
		for cell in col:
			lastNames.append(cell.value)
	fullNames = []
	# Combine first and last name to list of full names
	for i in range(len(firstNames)):
		logging.debug(f'\nfirstname = {firstNames[i]}\nlastNames = {lastNames[i]}')
		fullNames.append(firstNames[i]+' '+lastNames[i])
	# Write full names in row A
	c= 0 #cell counter
	for col in ws.iter_cols(min_row=2, min_col=1, max_row=maxim, max_col=1):
		for cell in col:
			cell.value = fullNames[c]
			logging.debug(f'cell value {cell.value}')
			c+=1

# Generates list of keys based on the first document.
def get_key_list (first_list):
	key_list = []
	for keys in first_list[0]:
		key_list.append(keys)
	return key_list

# Looks to find differences in the two datasets.
def compare_columns(first_list, second_list, key_list):
	with open('outputFile.txt','w') as outputFile: # Initiating the output file
		c=1 # Numeration
		for v in range(2): # Two datasets to compare
			for i in key_list: # The keys should be the same in both datasets
				try:
					if first_list[v][i] != second_list[v][i]:
						if v+1 == 1:
							colName = 'Residual'
						else:
							colName = 'Vacation days'
						print(f'{c}) Difference found:')
						outputFile.write('Difference found:\n')
						print(f'Values for key >> {i} << in column {colName} do not correspond.')
						outputFile.write(f'Values for key >> {i} << in column {colName} do not correspond.\n')
						print(f'\t{FIRST_LIST_DISPLAY_NAME} = {first_list[v][i]}\n\t{SECOND_LIST_DISPLAY_NAME} = {second_list[v][i]}\n')
						outputFile.write(f'\t{FIRST_LIST_DISPLAY_NAME} = {first_list[v][i]}\n\t{SECOND_LIST_DISPLAY_NAME} = {second_list[v][i]}\n\n')
						c+=1
				except KeyError:
					# Common KeyError if the datasets keys aren't the same
					print(f"\n{c}) DATA MISSING - The key >>> {i} <<< exists only in the first document.\nPlease ensure both documents have corresponding keys.\n")
					outputFile.write(f"\n{c}) The key >>> {i} <<< exists only in one document.\nPlease ensure both documents have corresponding keys.\n\n")
					c+=1

FIRST_FILE_ADDRESS = varDepo.variables[0]['FIRST_FILE_ADDRESS']
SECOND_FILE_ADDRESS = varDepo.variables[0]['SECOND_FILE_ADDRESS']
KEY_COLUMN = varDepo.variables[0]['KEY_COLUMN'] # Has to be same for both documents

# Both column letters need to be in the order of lookup.
FIRST_VALUE_COLUMN = varDepo.variables[0]['FIRST_VALUE_COLUMN']
SECOND_VALUE_COLUMN = varDepo.variables[0]['SECOND_VALUE_COLUMN']

FIRST_LIST_DISPLAY_NAME=varDepo.variables[0]['FIRST_LIST_DISPLAY_NAME']
SECOND_LIST_DISPLAY_NAME= varDepo.variables[0]['SECOND_LIST_DISPLAY_NAME']

# First doc
wb, ws = load_data(FIRST_FILE_ADDRESS, 'Employees') # Load the document.
format_names(ws, 1)	# Fix input errors for keys and standardise format.
wb.save(FIRST_FILE_ADDRESS) # Save shanges in the document.
first_list = create_dict_list(ws, KEY_COLUMN, FIRST_VALUE_COLUMN) # Extracts data in a list.

# Second doc
wb, ws = load_data(SECOND_FILE_ADDRESS, None) # Load the document.
construct_full_name(ws) # Form full name from 'First name' and 'Last name' columns.
format_names(ws, 1) # Fix input errors in keys and standardise format.
wb.save(SECOND_FILE_ADDRESS) # Save changes to the document.
second_list = create_dict_list(ws, KEY_COLUMN, SECOND_VALUE_COLUMN) # Extract data in a list.

# Analysis
key_list= get_key_list(first_list) # Generate a list of keys.
compare_columns(first_list, second_list, key_list) # Compare data sets and output findings.

# Delete xlsx docs.
send2trash.send2trash(FIRST_FILE_ADDRESS)
send2trash.send2trash(SECOND_FILE_ADDRESS)