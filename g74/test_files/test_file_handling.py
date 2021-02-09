"""
#######################################################################################################################
###											PSSE G74 Fault Studies													###
###		Unit tests associated with the processing and manipulation of files											###																													###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

import unittest
import os
import sys

import g74
import g74.file_handling as test_module
import pandas as pd

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_LOGS = os.path.join(TESTS_DIR, 'logs')

two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)

DELETE_LOG_FILES = True
# Set to True and any files created during the tests will also be deleted
DELETE_TEST_FILES = False


# ----- UNIT TESTS -----
class TestBusbarImport(unittest.TestCase):
	"""
		Tests that a spreadsheet of busbar numbers can successfully be imported
	"""
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestBusbarData', debug=g74.constants.DEBUG_MODE)
		cls.busbars_file = os.path.join(TESTS_DIR, 'test_busbars.xlsx')

	def test_import_busbars_success(self):
		"""
			Tests that a list of busbars can be successfully imported
		:return:
		"""
		list_of_busbars = test_module.import_busbars_list(path=self.busbars_file)
		self.assertEqual(list_of_busbars[0], 10)
		self.assertEqual(list_of_busbars[5], 100)
		self.assertTrue(len(list_of_busbars) == 8)

	def test_import_busbars_error(self):
		"""
			Tests that a list of busbars can be successfully imported but that some
			error messages are raised
		:return:
		"""
		list_of_busbars = test_module.import_busbars_list(
			path=self.busbars_file, sheet_number=1
		)
		self.assertEqual(list_of_busbars[0], 10)
		self.assertEqual(list_of_busbars[4], 100)
		self.assertTrue(len(list_of_busbars) == 7)

	@classmethod
	def tearDownClass(cls):
		# Delete log files created by logger
		if DELETE_LOG_FILES:
			paths = [
				cls.logger.pth_debug_log,
				cls.logger.pth_progress_log,
				cls.logger.pth_error_log
			]
			del cls.logger
			for pth in paths:
				if os.path.exists(pth):
					os.remove(pth)


class TestWorksheetNameChecker(unittest.TestCase):
	"""
		Tests the function to check worksheet names works correctly
	"""
	# Files are added to this as run and then deleted if no longer needed at the end
	files_to_delete = list()  # type: list

	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestWorksheetNameChecker', debug=g74.constants.DEBUG_MODE)

	def setUp(self):
		""" Same workbook is tested every time """
		self.test_wkbk = os.path.join(TESTS_DIR, 'test_wkbk.xlsx')
		self.files_to_delete.append(self.test_wkbk)

		# Confirm workbook doesn't already exist and if it does then delete
		if os.path.isfile(self.test_wkbk):
			os.remove(self.test_wkbk)

	def test_unique_name(self):
		""" Checks that if a unique name is passed it returns correctly """
		test_unique_name = 'TEST_SHEET_UNIQUE'

		with pd.ExcelWriter(path=self.test_wkbk, engine=test_module.excel_engine) as wkbk:
			new_name = test_module.worksheet_name_checker(wkbk=wkbk, sheet_name=test_unique_name)

		self.assertEqual(new_name, test_unique_name)

	def test_non_unique_name(self):
		""" Checks that if a non unique name is passed it returns correctly """
		test_name = 'TEST_SHEET_UNIQUE'

		with pd.ExcelWriter(path=self.test_wkbk, engine=test_module.excel_engine) as wkbk:
			wkbk.book.add_worksheet(name=test_name)
			new_name = test_module.worksheet_name_checker(wkbk=wkbk, sheet_name=test_name)

		# Confirm that names not matching any more
		self.assertNotEqual(new_name, test_name)
		expected_new_name = '{}(1)'.format(test_name)
		self.assertEqual(new_name, expected_new_name)

	def test_too_many_worksheets(self):
		""" Checks that if a non unique name is passed it returns correctly """
		test_name = 'TEST_SHEET_UNIQUE'

		with pd.ExcelWriter(path=self.test_wkbk, engine=test_module.excel_engine) as wkbk:
			wkbk.book.add_worksheet(name=test_name)
			for i in range(1, 105):
				name = '{}({})'.format(test_name, i)
				wkbk.book.add_worksheet(name=name)

			# Confirm stops at 100 iterations
			self.assertRaises(StopIteration, test_module.worksheet_name_checker, wkbk=wkbk, sheet_name=test_name)

	@classmethod
	def tearDownClass(cls):
		""" Deletes any files that are created """
		if DELETE_TEST_FILES:
			for f in cls.files_to_delete:
				if os.path.isfile(f):
					os.remove(f)

if __name__ == '__main__':
	unittest.main()
