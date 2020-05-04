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

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_LOGS = os.path.join(TESTS_DIR, 'logs')

two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)

DELETE_LOG_FILES = True


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


if __name__ == '__main__':
	unittest.main()
