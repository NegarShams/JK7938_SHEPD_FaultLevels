import unittest
import os
import sys
import g74
import g74.gui as test_module


TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_LOGS = os.path.join(TESTS_DIR, 'logs')

two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)

DELETE_LOG_FILES = True


class TestGui(unittest.TestCase):
	"""
		UnitTest package to confirm that GUI can be produced correctly
	"""
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Initialise logger
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestGUI', debug=g74.constants.DEBUG_MODE)

	def test_file_selection(self):
		"""Tests for the file selection function"""
		gui = test_module.MainGUI(title='Testing GUI')

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
				try:
					if os.path.exists(pth):
						os.remove(pth)
				except WindowsError:
					print('Unable to delete file: {}'.format(pth))
