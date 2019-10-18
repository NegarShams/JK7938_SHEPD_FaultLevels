"""
#######################################################################################################################
###											PSSE G7/4 Fault Studies													###
###		Script sets up PSSE to carry out fault studies in line with requirements of ENA G7/4						###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

import string
import logging
import pandas as pd
import g74.constants as constants


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
