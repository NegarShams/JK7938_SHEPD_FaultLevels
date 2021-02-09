"""
#######################################################################################################################
###											PSSE G74 Fault Studies													###
###		Script sets up PSSE to carry out fault studies in line with requirements of ENA G74							###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

import string
import logging
import pandas as pd
import xlsxwriter
import g74.constants as constants

# Engine to use for writing to excel, has to be XlsxWriter to ensure tab colours can be changed
excel_engine = 'xlsxwriter'


def colnum_string(n):
	test_str = ""
	while n > 0:
		n, remainder = divmod(n - 1, 26)
		test_str = chr(65 + remainder) + test_str
	return test_str


def colstring_number(col):
	num = 0
	for c in col:
		if c in string.ascii_letters:
			num = num * 26 + (ord(c.upper()) - ord('A')) + 1
	return num


def import_busbars_list(path, sheet_number=0):
	"""
		Imports all busbars listed in a file assuming they are the first column.
		TODO: Update to include some data processing to identify busbar numbers
	:param str path:  Full path of file to be imported
	:param int sheet_number:  Number of sheet to import
	:return list busbars:  List of busbars as integers
	"""
	logger = logging.getLogger(constants.Logging.logger_name)
	# Column number in DataFrame which will contain busbar numbers
	col_num = 0

	# Import excel workbook and then process
	df_busbars = pd.read_excel(io=path, sheet_name=sheet_number, header=None)
	logger.debug('Imported list of busbars from file: {}'.format(path))

	# Process imported DataFrame and convert all busbars to integers then report any which could not be converted
	busbars_series = df_busbars.iloc[:, col_num]
	# Try and convert all values to integer, if not possible then change to nan value
	busbars = pd.to_numeric(busbars_series, errors='coerce', downcast='integer')
	list_of_errors = busbars.isnull()

	# Check for any errors
	if list_of_errors.any():
		error_busbars = busbars_series[list_of_errors]
		msg0 = 'The following entries in the spreadsheet: {} could not be converted to busbar integers'.format(path)
		msg1 = '\n'.join(
			[
				'\t- Busbar <{}>'.format(bus) for bus in error_busbars
			])
		logger.error('{}\n{}'.format(msg0, msg1))
	else:
		logger.debug('All busbars successfully converted')

	list_of_busbars = list(busbars_series[~list_of_errors])
	return list_of_busbars


def worksheet_name_checker(wkbk, sheet_name):
	"""
		Function checks if a worksheet already exists in a workbook and if it does then instead returns a different name
		to use
	:param pd.ExcelWriter wkbk:  Instance of excel workbook to check
	:param str sheet_name:  Proposed sheet_name
	:return str sheet_name:  Resulting sheet_name
	"""
	logger = logging.getLogger(constants.Logging.logger_name)

	# Check if sheet already exists
	i = 0
	orig_sheet_name = sheet_name
	sheet_already_exists = sheet_name in wkbk.book.sheetnames.keys()

	# If sheet already exists then loop through until a name is found that doesn't clash
	while sheet_already_exists:
		# Iterate counter
		i += 1
		# Rename sheet to include (i) at the end
		sheet_name = '{}({})'.format(orig_sheet_name, i)
		sheet_already_exists = sheet_name in wkbk.book.sheetnames.keys()

		if i > 100:
			logger.critical(
				(
					'When trying to find a unique worksheet with name {} in workbook {} reached {} iterations which '
					'suggests either {} worksheets with the name {} already exist or there is some other error.'
				).format(sheet_name, wkbk.path, i, sheet_name, i-1)
			)
			raise StopIteration('Very high number of iterations to stopping, check error message above')

	if orig_sheet_name != sheet_name:
		logger.warning(
			(
				'The worksheet named {} already exists in the excel workbook {} and rather than being overwritten a new '
				'sheet will be created with the name {}'
			).format(orig_sheet_name, wkbk.path, sheet_name)
		)

	return sheet_name


def write_fault_data_to_excel(pth, df, message, sheet_name, tab_color=None):
	"""
		Function deals with writing a DataFrame to excel for the fault current data and ensures a message
		is added along with the correct sheet name and location
	:param str pth:  Full path to excel workbook to write
	:param pd.DataFrame df:  Pandas Dataframe to write
	:param str message:  Message to include on first row
	:param str sheet_name:  Name of sheet to use (an additional sheet is also created with the name transposed)
	:param str tab_color:  Hexidemical code for tab_color to use
	:return None:
	"""
	logger = logging.getLogger(constants.Logging.logger_name)

	# Load workbook
	with pd.ExcelWriter(path=pth, engine=excel_engine) as wkbk:
		# Confirm sheet name isn't duplicated and then create new sheet
		sheet_name = worksheet_name_checker(wkbk=wkbk, sheet_name=sheet_name)
		wksh = wkbk.book.add_worksheet(name=sheet_name)

		# Create name for transposed sheet as well
		sheet_name_transposed = worksheet_name_checker(wkbk=wkbk, sheet_name='{}_transposed'.format(sheet_name))
		wksh_t = wkbk.book.add_worksheet(name=sheet_name_transposed)

		# Have to add worksheet to Pandas list of worksheets
		# (https://stackoverflow.com/questions/32957441/putting-many-python-pandas-dataframes-to-one-excel-worksheet)
		wkbk.sheets[sheet_name] = wksh
		logger.debug('New worksheet named {} added to workbook {}'.format(sheet_name, wkbk.path))
		wkbk.sheets[sheet_name_transposed] = wksh_t
		logger.debug('New worksheet named {} added to workbook {}'.format(sheet_name_transposed, wkbk.path))

		# Write some details on the status first and colour the tab accordingly
		row = 0
		col = 0
		wksh.write_string(row=row, col=col, string=message)
		wksh_t.write_string(row=row, col=col, string=message)
		# Only set tab_color if not None
		if tab_color:
			wksh.set_tab_color(tab_color)
			wksh_t.set_tab_color(tab_color)

		# Write DataFrame to excel worksheet
		df.to_excel(wkbk, sheet_name=sheet_name, startrow=row+constants.Excel.row_spacing)
		logger.debug('DataFrame written to worksheet {}'.format(sheet_name))
		df.T.to_excel(wkbk, sheet_name=sheet_name_transposed, startrow=row+constants.Excel.row_spacing)
		logger.debug('Transposed DataFrame written to worksheet {}'.format(sheet_name_transposed))

	return None
