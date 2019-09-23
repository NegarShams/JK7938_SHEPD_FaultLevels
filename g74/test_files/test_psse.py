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

import unittest
import os
import sys
import g74
import g74.psse as test_module

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_LOGS = os.path.join(TESTS_DIR, 'logs')

SAV_CASE_COMPLETE = os.path.join(TESTS_DIR, 'JK7938_SAV_TEST.sav')

two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)

SKIP_SLOW_TESTS = False
DELETE_LOG_FILES = True

# These constants are used to return the environment back to its original format
# for testing the PSSPY import functions
original_sys = sys.path
original_environ = os.environ['PATH']


# ----- UNIT TESTS -----
@unittest.skipIf(SKIP_SLOW_TESTS, 'PSSPY import testing skipped since slow to run')
class TestPsseInitialise(unittest.TestCase):
	"""
		Functions to check that PSSE import and initialisation is possible
	"""
	def test_psse32_psspy_import_fail(self):
		"""
			Test that PSSE version 32 cannot be initialised because it is not installed
		:return:
		"""
		sys.path = original_sys
		os.environ['PATH'] = original_environ
		self.psse = test_module.InitialisePsspy(psse_version=32)
		self.assertIsNone(self.psse.psspy)

	def test_psse33_psspy_import_success(self):
		"""
			Test that PSSE version 33 can be initialised
		:return:
		"""
		self.psse = test_module.InitialisePsspy(psse_version=33)
		self.assertIsNotNone(self.psse.psspy)

	def test_psse34_psspy_import_success(self):
		"""
			Test that PSSE version 34 can be initialised
		:return:
		"""
		self.psse = test_module.InitialisePsspy(psse_version=34)
		self.assertIsNotNone(self.psse.psspy)

		# Initialise psse
		status = self.psse.initialise_psse()
		self.assertTrue(status)

	def tearDown(self):
		"""
			Tidy up by removing variables and paths that are not necessary
		:return:
		"""
		sys.path.remove(self.psse.psse_py_path)
		os.environ['PATH'] = os.environ['PATH'].strip(self.psse.psse_os_path)
		os.environ['PATH'] = os.environ['PATH'].strip(self.psse.psse_py_path)


@unittest.skipIf(SKIP_SLOW_TESTS, 'PSSE initialisation skipped since slow to run')
class TestPsseControl(unittest.TestCase):
	"""
		Unit test for loading of SAV case file and subsequent operations
	"""
	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestPsseControl', debug=g74.constants.DEBUG_MODE)
		cls.psse = test_module.PsseControl()
		cls.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

	def test_load_case_error(self):
		psse = test_module.PsseControl()
		with self.assertRaises(ValueError):
			psse.load_data_case()

	def test_load_flow_full(self):
		load_flow_success, df = self.psse.run_load_flow()
		self.assertTrue(load_flow_success)
		self.assertTrue(df.empty)

	def test_load_flow_flatstart(self):
		load_flow_success, df = self.psse.run_load_flow(flat_start=True)
		self.assertTrue(load_flow_success)
		self.assertTrue(df.empty)

	def test_load_flow_locked_taps(self):
		load_flow_success, df = self.psse.run_load_flow(lock_taps=True)
		self.assertTrue(load_flow_success)
		self.assertTrue(df.empty)

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
