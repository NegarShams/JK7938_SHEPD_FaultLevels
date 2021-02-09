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

import g74
import g74.psse as test_module
import g74.constants as constants

import unittest
import os
import sys
import pandas as pd
import numpy as np
import math

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_LOGS = os.path.join(TESTS_DIR, 'logs')

SAV_CASE_COMPLETE = os.path.join(TESTS_DIR, 'JK7938_SAV_TEST.sav')
SAV_CASE_COMPLETE2 = os.path.join(TESTS_DIR, 'JK7938_SAV_TEST2.sav')

two_up = os.path.abspath(os.path.join(TESTS_DIR, '../..'))
sys.path.append(two_up)

DELETE_LOG_FILES = True

# These constants are used to return the environment back to its original format
# for testing the PSSPY import functions
original_sys = sys.path
original_environ = os.environ['PATH']


# ----- UNIT TESTS -----
class TestPsseInitialise(unittest.TestCase):
	"""
		Functions to check that PSSE import and initialisation is possible
	"""
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestPSSEInitialise', debug=g74.constants.DEBUG_MODE)

	def test_psse33_psspy_import_success(self):
		"""
			Test that PSSE version 33 can be initialised
		:return:
		"""
		sys.path = original_sys
		os.environ['PATH'] = original_environ
		self.psse = test_module.InitialisePsspy(psse_version=33)
		self.assertIsNotNone(self.psse.psspy)

	def test_psse34_psspy_import_success(self):
		"""
			Test that PSSE version 34 can be initialised
		:return:
		"""
		sys.path = original_sys
		os.environ['PATH'] = original_environ
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
		sys.path = original_sys
		os.environ['PATH'] = original_environ

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


class TestPsseControl(unittest.TestCase):
	"""
		Unit test for loading of SAV case file and subsequent operations
	"""
	# To avoid error message
	logger = None

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

	def test_bus_subsystem(self):
		""" Tests the bus subsystem can be defined """
		sid = self.psse.define_bus_subsystem(buses=[11, 33], sid=2)
		self.assertEqual(sid, self.psse.sid)
		self.psse.sid = -1

	def test_bus_subsystem_fails(self):
		""" Tests the bus subsystem can be defined """
		constants.PSSE.sid = 20
		self.assertRaises(ValueError, self.psse.define_bus_subsystem, [11, 33], 25)
		self.assertEqual(-1, self.psse.sid)
		constants.PSSE.sid = 1

	def test_bus_subsystem_fails2(self):
		""" Tests the bus subsystem defines based on the constants value """
		sid = self.psse.define_bus_subsystem(buses=[11, 33], sid=25)
		self.assertEqual(constants.PSSE.sid, sid)
		self.psse.sid = -1

	def test_bus_subsystem_no_buses(self):
		sid = self.psse.define_bus_subsystem(buses=[], sid=2)
		self.assertEqual(-1, sid)

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


class TestBkdyNonPSSEComponents(unittest.TestCase):
	"""
		Unit test for individual components required for BKDY calculation that do not need PSSE
	"""
	# Reference added here to avoid issue with referencing in tearDownClass method
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestBkdyComponents', debug=g74.constants.DEBUG_MODE)
		# #cls.bkdy = test_module.BkdyFaultStudy()
		cls.output_file = os.path.join(TESTS_DIR, 'bkdy_output{}'.format(constants.General.ext_csv))

	def test_bkdy_file_import(self):
		"""
			Tests the reading of the BKDY export into the appropriate format
		"""
		bkdy_file = test_module.BkdyFile(output_file=self.output_file, fault_time=0.01)

		df = bkdy_file.process_bkdy_output()
		self.assertAlmostEqual(df.loc[1, constants.BkdyFileOutput.ibasym], 1.6297, places=2)
		self.assertAlmostEqual(df.loc[1501, constants.BkdyFileOutput.ip], 3.9073, places=2)
		self.assertAlmostEqual(df.loc[5001, constants.BkdyFileOutput.ik11], 6.1801, places=2)
		self.assertAlmostEqual(df.loc[5101, constants.BkdyFileOutput.ibsym], 3.7529, places=2)

	def test_bkdy_file_import_fails(self):
		"""
			Checks that if a file has been deleted and attempts to process again then an error
			is shown and if the DataFrame is already empty then raises a SyntaxError
		"""
		bkdy_file = test_module.BkdyFile(output_file=self.output_file, fault_time=0.01)

		# Remove the reference to the file before it is processed which should result in an error
		# being raised when trying to process the file.
		bkdy_file.output_file = None
		self.assertRaises(SyntaxError, bkdy_file.process_bkdy_output)

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


class TestBkdyComponents(unittest.TestCase):
	"""
		Unit test for individual components required for BKDY calculation
	"""
	# Reference added here to avoid issue with referencing in tearDownClass method
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestBkdyComponents', debug=g74.constants.DEBUG_MODE)
		cls.psse = test_module.PsseControl()
		cls.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		cls.bkdy = test_module.BkdyFaultStudy(psse_control=cls.psse)
		cls.output_file = os.path.join(TESTS_DIR, 'bkdy_output{}'.format(constants.General.ext_csv))

	def test_convert(self):
		""" Test converting of generators """
		# Convert SAV case and check status flag is set to True
		self.psse.convert_sav_case()
		self.assertTrue(self.psse.converted)

		# Reload sav case to avoid staying in converted format
		self.psse.load_data_case()
		self.assertFalse(self.psse.converted)

	def test_machine_idev(self):
		"""
			Tests the production of the idev file needed for the machines
		"""
		# IDEV file
		machine_idev_file = os.path.join(TESTS_DIR, 'machines{}'.format(constants.PSSE.ext_bkd))
		if os.path.exists(machine_idev_file):
			os.remove(machine_idev_file)

		mac_data = test_module.MachineData()
		mac_data.produce_idev(target=machine_idev_file)

		self.assertTrue(os.path.exists(machine_idev_file))
		os.remove(machine_idev_file)

	def test_induction_idev(self):
		"""
			Tests the production of the idev file needed for the machines
		"""
		# IDEV file
		idev_file = os.path.join(TESTS_DIR, 'induction{}'.format(constants.PSSE.ext_bkd))
		if os.path.exists(idev_file):
			os.remove(idev_file)

		mac_data = test_module.InductionData()
		mac_data.add_to_idev(target=idev_file)

		self.assertTrue(os.path.exists(idev_file))
		os.remove(idev_file)

	def test_complete_idev(self):
		"""
			Tests the production of the idev file needed for the machines
		"""
		# IDEV file
		idev_file = os.path.join(TESTS_DIR, 'impedances{}'.format(constants.PSSE.ext_bkd))
		if os.path.exists(idev_file):
			os.remove(idev_file)

		mac_data = test_module.MachineData()
		mac_data.produce_idev(target=idev_file)
		induction_machines = test_module.InductionData()
		induction_machines.add_to_idev(target=idev_file)

		self.assertTrue(os.path.exists(idev_file))
		# #os.remove(idev_file)

	def test_bkdy_calc(self):
		"""
			Carries out the BKDY calculation method and checks output file produced
		"""
		# IDEV file
		idev_file = os.path.join(TESTS_DIR, 'test{}'.format(constants.PSSE.ext_bkd))
		if os.path.exists(idev_file):
			os.remove(idev_file)
		# Output file for BKDY is defined above
		if os.path.exists(self.output_file):
			os.remove(self.output_file)

		self.bkdy.create_breaker_duty_file(target_path=idev_file)
		self.bkdy.main(output_file=self.output_file, fault_time=0.06, name='0.06')

		self.assertTrue(os.path.exists(idev_file))
		self.assertTrue(os.path.exists(self.output_file))
		os.remove(idev_file)

	def test_bkdy_calc_all(self):
		"""
			Carries out the BKDY calculation method and checks output file produced
		"""
		# IDEV file
		idev_file = os.path.join(TESTS_DIR, 'test{}'.format(constants.PSSE.ext_bkd))
		output_peak = os.path.join(TESTS_DIR, 'peak_currents{}'.format(constants.General.ext_csv))
		output_break = os.path.join(TESTS_DIR, 'break_currents{}'.format(constants.General.ext_csv))
		if os.path.exists(idev_file):
			os.remove(idev_file)
		# Output file for BKDY is defined above
		for x in (output_break, output_peak):
			if os.path.exists(x):
				os.remove(x)

		self.bkdy.create_breaker_duty_file(target_path=idev_file)
		self.bkdy.main(output_file=output_peak, fault_time=0.01, name='0.01')
		self.bkdy.main(output_file=output_break, fault_time=0.06, name='0.06')

		# Check files created and then remove
		for path in (output_break, output_peak, idev_file):
			self.assertTrue(os.path.exists(path))
			os.remove(path)

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


class TestIECComponents(unittest.TestCase):
	"""
		Tests the calculation of fault currents using the IEC method
	"""
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestIECcomponents', debug=g74.constants.DEBUG_MODE)
		cls.psse = test_module.PsseControl()
		cls.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

	def test_calculate_iec_3ph(self):
		"""
			Function tests complete calculation of IEC faults using 3 phase fault current calculation
		:return:
		"""
		# Reload SAV case
		self.psse.load_data_case()

		# Bus numbers to test - These are busbars with motors directly contributing to a fault at
		# this busbar
		buses_to_test = [11, 33]
		fault_times_to_test = (0.05, 0.06)

		# Model setup
		# Add machines to model
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		# IEC method for fault current calculations
		iec = test_module.IecFaults(psse=self.psse, buses=buses_to_test)

		# Iterative loop that tests a range of fault times
		dfs = list()

		# Run calculation for all fault times
		for fault_time in fault_times_to_test:

			# Calculate new machine impedance values
			g74_data.calculate_machine_impedance(fault_time=fault_time)

			# Add machines to model
			g74_data.add_machines()

			if fault_time == constants.PSSE.min_fault_time:
				self.psse.save_data_case(pth_sav=SAV_CASE_COMPLETE2)

			# Run fault study for both busbars at this time step
			df = iec.fault_study(fault_time=fault_time, lll=True)
			dfs.append(df)

	def test_calculate_iec_lg(self):
		"""
			Function tests complete calculation of IEC faults using 3 phase fault current calculation
		:return:
		"""
		# Reload SAV case
		self.psse.load_data_case()

		# Bus numbers to test - These are busbars with motors directly contributing to a fault at
		# this busbar
		buses_to_test = [11, 33]
		fault_times_to_test = (0.05, 0.06)

		# Model setup
		# Add machines to model
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		# IEC method for fault current calculations
		iec = test_module.IecFaults(psse=self.psse, buses=buses_to_test)

		# Iterative loop that tests a range of fault times
		dfs = list()

		# Run calculation for all fault times
		for fault_time in fault_times_to_test:

			# Calculate new machine impedance values
			g74_data.calculate_machine_impedance(fault_time=fault_time)

			# Add machines to model
			g74_data.add_machines()

			if fault_time == constants.PSSE.min_fault_time:
				self.psse.save_data_case(pth_sav=SAV_CASE_COMPLETE2)

			# Run fault study for both busbars at this time step
			df = iec.fault_study(fault_time=fault_time, lll=False, lg=True)
			dfs.append(df)

	def test_calculate_trying_to_calculate_two_faults_fails(self):
		"""
			Function tests calculation of LLL and LG fault times fails
		:return:
		"""
		# Reload SAV case
		self.psse.load_data_case()

		# Bus numbers to test - These are busbars with motors directly contributing to a fault at
		# this busbar
		buses_to_test = [11, 33]
		fault_time = 0.05

		# Model setup
		iec = test_module.IecFaults(psse=self.psse, buses=buses_to_test)

		# Run fault study for both lll and lg and should raise a syntax error
		self.assertRaises(SyntaxError, iec.fault_study, fault_time=fault_time, lll=True, lg=True)

	def test_calculate_trying_to_calculate_no_faults_fails(self):
		"""
			Function tests calculation of LLL and LG fault times fails
		:return:
		"""
		# Reload SAV case
		self.psse.load_data_case()

		# Bus numbers to test - These are busbars with motors directly contributing to a fault at
		# this busbar
		buses_to_test = [11, 33]
		fault_time = 0.05

		# Model setup
		iec = test_module.IecFaults(psse=self.psse, buses=buses_to_test)

		# Run fault study for both lll and lg and should raise a syntax error
		self.assertRaises(ValueError, iec.fault_study, fault_time=fault_time, lll=False, lg=False)

	def test_calculate_iec_method_lll(self):
		"""
			Confirms that the fault current contribution using the IEC method is calculated correctly for a 3 phase
			fault

		"""
		# File constants
		fault_times = list(np.arange(0.0001, 0.12, 0.01))
		buses_to_fault = [1102, 3302]

		# Load model
		self.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		# Model setup
		# Add machines to model
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		# IEC method for fault current calculations
		iec = test_module.IecFaults(psse=self.psse, buses=buses_to_fault)

		df = iec.calculate_fault_currents(
			fault_times=fault_times, g74_infeed=g74_data,
			lll=True
		)

		# Confirm returned DataFrame is the expected shape
		# TODO: Confirm some of the returned values are correct
		self.assertEqual(df.shape, (2, 55))

		pass

	def test_extract_lll_fault_results_example(self):
		"""
			Confirms that the fault current contribution is extracted correctly
		"""
		# Test export file
		test_psse_export = os.path.join(TESTS_DIR, 'iec_3ph_export.xlsx')
		# Confirm doesn't exist and if it does then delete
		if os.path.isfile(test_psse_export):
			os.remove(test_psse_export)

		# File constants
		fault_times = [0.05]

		# Load model
		self.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		# Model setup
		# Add machines to model
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		# IEC method for fault current calculations
		iec = test_module.IecFaults(psse=self.psse)

		df = iec.calculate_fault_currents(
			fault_times=fault_times, g74_infeed=g74_data,
			lll=True
		)

		# Write to excel file using writer as part of file_handling
		g74.file_handling.write_fault_data_to_excel(
			pth=test_psse_export, df=df, message='FAULT CURRENT TEST OUTPUT',
			sheet_name=g74.constants.Excel.iec_sheet_name_lll,
			tab_color=g74.constants.Excel.iec_tab_color
		)

		# Save final PSSE model for further investigation
		self.psse.save_data_case(pth_sav=os.path.join(TESTS_DIR, 'IEC_SAV_EXTRACT.sav'))

		# Confirm file exists
		self.assertTrue(os.path.isfile(test_psse_export))

	def test_calculate_iec_method_lg(self):
		"""
			Confirms that the fault current contribution using the IEC method is calculated correctly for LG faults

		"""
		# File constants
		fault_times = list(np.arange(0.0001, 0.12, 0.01))
		buses_to_fault = [1102, 3302]

		# Load model
		self.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		# Model setup
		# Add machines to model
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		# IEC method for fault current calculations
		iec = test_module.IecFaults(psse=self.psse, buses=buses_to_fault)

		df = iec.calculate_fault_currents(
			fault_times=fault_times, g74_infeed=g74_data,
			lll=False, lg=True
		)

		# Confirm returned DataFrame is the expected shape
		# TODO: Confirm some of the returned values are correct
		self.assertEqual(df.shape, (2, 55))

		pass

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


class TestBkdyIntegration(unittest.TestCase):
	"""
		Unit test for individual components required for BKDY calculation
	"""
	# Reference added here to avoid issue with referencing in tearDownClass method
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestBkdyComponents', debug=g74.constants.DEBUG_MODE)
		cls.psse = test_module.PsseControl()
		cls.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

	def test_bkdy_calculation(self):
		"""
			Test complete calculation works
		:return:
		"""
		# IDEV file
		idev_file = os.path.join(TESTS_DIR, 'test{}'.format(constants.PSSE.ext_bkd))
		output_peak = os.path.join(TESTS_DIR, 'peak_currents{}'.format(constants.General.ext_csv))
		output_break = os.path.join(TESTS_DIR, 'break_currents{}'.format(constants.General.ext_csv))

		# Paths and fault times to be considered
		output_paths = (output_peak, output_break)
		fault_times = (0.01, 0.06)

		bkdy = test_module.BkdyFaultStudy(psse_control=self.psse)
		bkdy.create_breaker_duty_file(target_path=idev_file)

		# Run fault study on each fault time
		for path, fault_time in zip(output_paths, fault_times):
			bkdy.main(output_file=path, fault_time=fault_time, name='{:.2f}'.format(fault_time))

		# Process results of BKDY files into relevant inputs
		dfs = list()
		for path, fault_time in zip(output_paths, fault_times):
			# File manually set for unittesting
			bkdy_file = test_module.BkdyFile(output_file=path, fault_time=fault_time)
			df = bkdy_file.process_bkdy_output(delete=True)
			# Delete original file once processed
			dfs.append(df)

		# Delete idev file
		os.remove(idev_file)

	def test_bkdy_g74_method(self):
		"""
			Test that bkdy calculation works with contribution from
			LV connected machines
		"""
		# File constants
		idev_file = os.path.join(TESTS_DIR, 'test{}'.format(constants.PSSE.ext_bkd))
		output_peak = os.path.join(TESTS_DIR, 'peak_currents{}'.format(constants.General.ext_csv))
		output_break = os.path.join(TESTS_DIR, 'break_currents{}'.format(constants.General.ext_csv))
		excel_export = os.path.join(TESTS_DIR, 'excel_export_g74{}'.format('.xlsx'))

		output_files = (output_peak, output_break)
		fault_times = (0.01, 0.06)
		names = (constants.SHEPD.cb_make, constants.SHEPD.cb_break)

		# Add contribution from embedded machines
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()
		g74_data.add_machines()

		# Carry out BKDY calculation
		bkdy = test_module.BkdyFaultStudy(psse_control=self.psse)
		bkdy.create_breaker_duty_file(target_path=idev_file)
		# Run fault study for each fault time
		for f, flt_time, name in zip(output_files, fault_times, names):
			bkdy.main(output_file=f, fault_time=flt_time, name=name)

		# Process output files into DataFrames and export to Excel
		df = bkdy.combine_bkdy_output()

		# Check if any of the paths already exist and if they do delete them
		for path in (excel_export, SAV_CASE_COMPLETE2):
			if os.path.exists(path):
				os.remove(path)

		# Test exporting and saving results
		df.to_excel(excel_export)
		self.psse.save_data_case(pth_sav=SAV_CASE_COMPLETE2)

		# Confirm newly created files exist and then delete
		for path in (excel_export, SAV_CASE_COMPLETE2, idev_file):
			self.assertTrue(os.path.exists(path))
			os.remove(path)

	def test_calculate_machine_time_dependant_contribution_bkdy(self):
		"""
			Confirms that the fault current contribution from the embedded machines
			is correct on both sides of the transformer.

			Note: This has to be carried out using BKDY since the IEC method includes
			a correction factor for the transformer.
		"""
		# File constants
		idev_file = os.path.join(TESTS_DIR, 'test{}'.format(constants.PSSE.ext_bkd))
		fault_times = list(np.arange(0.0, 0.12, 0.01))
		buses_to_fault = [1102, 3302]
		excel_export = os.path.join(TESTS_DIR, 'excel_export_g74{}'.format('.xlsx'))

		# Parameters to use for checking calculation
		pre_fault_v = 1.0
		# Load value modelled as 5.0 MVA in PSSe test model
		load_value = 5.0

		# Load model
		self.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		# Create circuit breaker duty file
		bkdy = g74.psse.BkdyFaultStudy(psse_control=self.psse)
		bkdy.create_breaker_duty_file(target_path=idev_file)

		# Switch off slack machine connect to 1102 so doesn't contribute to fault current for testing
		self.psse.switch_machine(i=1102, id='1')

		# Update model to include contribution from embedded machines
		g74_data = g74.psse.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		df = bkdy.calculate_fault_currents(
			fault_times=fault_times, g74_infeed=g74_data,
			buses=buses_to_fault, delete=False
		)

		# Calculate expected values for each fault time
		# Iterative loop that tests a range of fault times
		dfs = list()
		expected_f33_infeed_pu = list()
		expected_f11_infeed_pu = list()

		# R values remain constant for all fault times
		new_r33_value = constants.G74.rpos
		new_r11_value = new_r33_value - constants.G74.tx_r
		for fault_time in fault_times:
			# Calculate the expected X value and expected fault in feed in per unit
			# Target busbar is 11kV so need to account for 11/33kV transformer impedance
			if fault_time > constants.PSSE.min_fault_time:
				new_x33_value = 1.0 / ((1.0 / constants.G74.x11) * math.exp(-fault_time / constants.G74.t11))
			else:
				new_x33_value = constants.G74.x11

			new_x11_value = new_x33_value - constants.G74.tx_x

			expected_f33_infeed_pu.append(pre_fault_v/(new_r33_value**2 + new_x33_value**2)**0.5)
			expected_f11_infeed_pu.append(pre_fault_v/(new_r11_value**2 + new_x11_value**2)**0.5)

		# Calculate the values that would be expected based on the fault times tested
		expected_values_11 = dict(zip(
			fault_times,
			[x*load_value*constants.G74.mva_11/(math.sqrt(3)*11.0) for x in expected_f11_infeed_pu]
		))
		expected_values_33 = dict(zip(
			fault_times,
			[x*load_value*constants.G74.mva_33/(math.sqrt(3)*33.0) for x in expected_f33_infeed_pu]
		))

		# Check if any of the paths already exist and if they do delete them
		for path in (excel_export, SAV_CASE_COMPLETE2):
			if os.path.exists(path):
				os.remove(path)

		# Test exporting and saving results
		df.to_excel(excel_export)
		self.psse.save_data_case(pth_sav=SAV_CASE_COMPLETE2)

		# Confirm newly created files exist and then delete
		for path in (SAV_CASE_COMPLETE2, idev_file):
			self.assertTrue(os.path.exists(path))
			os.remove(path)

	def test_bkdy_3ph_export(self):
		"""
			Carries out a fault study using BKDY method for a 3phase fault and extract to a spreadsheet
		"""
		# File to save exported results to
		test_psse_export = os.path.join(TESTS_DIR, 'bkdy_3ph_export.xlsx')

		# File constants
		idev_file = os.path.join(TESTS_DIR, 'test{}'.format(constants.PSSE.ext_bkd))
		fault_times = [0.05]

		# Confirm file doesn't already exist and if it does then delete it
		if os.path.isfile(test_psse_export):
			os.remove(test_psse_export)

		# Load model
		self.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

		# Create circuit breaker duty file
		bkdy = g74.psse.BkdyFaultStudy(psse_control=self.psse)
		bkdy.create_breaker_duty_file(target_path=idev_file)

		# Update model to include contribution from embedded machines
		g74_data = g74.psse.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		df = bkdy.calculate_fault_currents(
			fault_times=fault_times, g74_infeed=g74_data, delete=True
		)

		# Write to excel file using writer
		g74.file_handling.write_fault_data_to_excel(
			pth=test_psse_export, df=df, message='FAULT CURRENT TEST OUTPUT',
			sheet_name=g74.constants.Excel.bkdy_sheet_name,
			tab_color=g74.constants.Excel.bkdy_tab_color
		)

		# Confirm file exists
		self.assertTrue(os.path.isfile(test_psse_export))

	def test_infinity_in_results_handled_correctly(self):
		"""
			Routine tests that for results where the thevenin impedance is negative return either ***** or infinity
			this confirms that they are handled correctly.  For the test model this has been setup for faults on the
			1102 busbar
		:return: None
		"""
		# Reload the PSSE SAV case
		self.psse.load_data_case()

		# IDEV file
		idev_file = os.path.join(TESTS_DIR, 'test{}'.format(constants.PSSE.ext_bkd))
		output_currents1 = os.path.join(TESTS_DIR, 'infinity_currents1{}'.format(constants.General.ext_csv))
		output_currents2 = os.path.join(TESTS_DIR, 'infinity_currents2{}'.format(constants.General.ext_csv))

		# Busbar 1301 has been setup to return infinite values and -X for long fault durations
		bus = 1102
		# Fault times being tested
		ft1 = 0.06
		ft2 = 0.5

		# Paths and fault times to be considered
		output_paths = (output_currents1, output_currents2)
		fault_times = (ft1, ft2)

		bkdy = test_module.BkdyFaultStudy(psse_control=self.psse)
		bkdy.create_breaker_duty_file(target_path=idev_file)

		# Run fault study on each fault time
		for path, fault_time in zip(output_paths, fault_times):
			bkdy.main(output_file=path, fault_time=fault_time, name='{:.2f}'.format(fault_time))

		# Process results of BKDY files into relevant inputs
		dfs = dict()
		for path, fault_time in zip(output_paths, fault_times):
			# Confirm that either **** or Infinity in each of the files
			test_term_exists = False
			with open(path) as test_file:
				contents = test_file.read()
				if (
						constants.BkdyFileOutput.nan_term1 in contents or
						constants.BkdyFileOutput.nan_term2 in contents or
						constants.BkdyFileOutput.nan_term3 in contents
				):
					test_term_exists = True
			self.assertTrue(test_term_exists)

			# File manually set for unittesting
			bkdy_file = test_module.BkdyFile(output_file=path, fault_time=fault_time)
			df = bkdy_file.process_bkdy_output(delete=False)

			# Delete original file once processed
			dfs[fault_time] = df

		# Check that the calculated values in df are correct when dealing with the **** and nan values in the
		# results file
		df = pd.concat(dfs.values(), axis=1, keys=dfs.keys())
		# DC value from method 1 should be as per BKDY output file and method 2 converted to 0.0
		self.assertAlmostEqual(df.loc[bus, (ft2, constants.BkdyFileOutput.idc_method1)], 0.00000, places=3)
		self.assertAlmostEqual(df.loc[bus, (ft2, constants.BkdyFileOutput.idc_method2)], 0.00730, places=3)
		self.assertAlmostEqual(df.loc[bus, (ft1, constants.BkdyFileOutput.idc_method1)], 0.00000, places=3)
		self.assertAlmostEqual(df.loc[bus, (ft1, constants.BkdyFileOutput.idc_method2)], 0.23100, places=3)

		# Peak value from method 1 should be as per BKDY output file and method 2 converted to 0.0
		self.assertAlmostEqual(df.loc[bus, (ft2, constants.BkdyFileOutput.ip_method1)], 0.00570, places=3)
		self.assertAlmostEqual(df.loc[bus, (ft2, constants.BkdyFileOutput.ip_method2)], 0.01300, places=3)
		self.assertAlmostEqual(df.loc[bus, (ft1, constants.BkdyFileOutput.ip_method1)], 0.22440, places=3)
		self.assertAlmostEqual(df.loc[bus, (ft1, constants.BkdyFileOutput.ip_method2)], 0.45530, places=3)

		# X values for busbar 1301 should be negative
		self.assertTrue(df.loc[bus, (ft1, constants.BkdyFileOutput.x)] == 0.0)
		self.assertTrue(df.loc[bus, (ft2, constants.BkdyFileOutput.x)] == 0.0)

		# Delete idev file
		os.remove(idev_file)

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


class TestPsseLoadData(unittest.TestCase):
	"""
		Unit tests for the extraction of load data, calculation of equivalent machines and adding them to the PSSE
		model for the G74 fault contribution
	"""
	# Reference added here to avoid issue with referencing in tearDownClass method
	logger = None

	@classmethod
	def setUpClass(cls):
		"""
			Load the SAV case into PSSE for further testing
		"""
		# Initialise logger
		cls.logger = g74.Logger(pth_logs=TEST_LOGS, uid='TestBkdyComponents', debug=g74.constants.DEBUG_MODE)
		# #cls.bkdy = test_module.BkdyFaultStudy()
		cls.output_file = os.path.join(TESTS_DIR, 'bkdy_output{}'.format(constants.General.ext_csv))
		cls.psse = test_module.PsseControl()
		cls.psse.load_data_case(pth_sav=SAV_CASE_COMPLETE)

	def test_load_data_summary(self):
		"""
			Checks that able to obtain summary of load connected at each busbar
		"""
		load_data = test_module.LoadData()

		# Check DataFrame is not empty
		self.assertFalse(load_data.df.empty)

		df = load_data.summary()
		# Confirm that the summarised values are less than the total values
		self.assertTrue(len(df) < len(load_data.df))

		# Confirm particular values at both single load buses and multiple load buses
		self.assertAlmostEqual(df.loc[100, constants.Loads.load], 5.0249, places=3)
		self.assertAlmostEqual(df.loc[2501, constants.Loads.load], 5.6932, places=3)

	def test_bus_data(self):
		"""
			Confirm able to pull out busbar data with nominal voltage in the format expected
		"""
		bus_data = test_module.BusData()

		# Check DataFrame is not empty (if it is then bus_data.update() has not been run by default)
		self.assertFalse(bus_data.df.empty)

		self.assertAlmostEqual(bus_data.df.loc[100, constants.Busbars.nominal], 132.0, places=1)
		self.assertAlmostEqual(bus_data.df.loc[101, constants.Busbars.nominal], 33.0, places=1)
		self.assertAlmostEqual(bus_data.df.loc[2601, constants.Busbars.nominal], 11.0, places=1)

	def test_identifying_machines_lv_only(self):
		"""
			Confirm that machines identified at the relevant busbars as either HV or LV when no HV connected machines
			provided as an input
		"""
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()

		# The MVA values should be equal since there are no HV motors
		df = g74_data.df_machines
		self.assertTrue(df[constants.Loads.load].equals(df[constants.G74.label_mva]/constants.G74.mva_11))

	def test_identifying_machines_hv_only(self):
		"""
			Confirm that machines identified at the relevant busbars as either HV or LV when no HV connected machines
			provided as an input
		"""
		# Produce a DataFrame of busbar numbers for HV machines
		hv_machines = test_module.BusData().df

		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters(hv_machines=hv_machines)

		# The MVA values should be equal since there are no HV motors
		df = g74_data.df_machines
		self.assertTrue(df[constants.G74.label_mva].equals(df[constants.Loads.load]*constants.G74.mva_hv))

	def test_calculating_machine_parameters(self):
		"""
			Confirm that machine parameters correctly added to DataFrame
		"""
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		# The MVA values should be equal since there are no HV motors
		df = g74_data.df_machines

		# Confirm value is correct for 33kV
		self.assertAlmostEqual(df.loc[33, constants.Machines.xsubtr], constants.G74.x11, places=4)
		# Confirm value is correct for 11kV
		# TODO: Will need to be revised if improved calculation for 11kV equivalents
		self.assertAlmostEqual(
			df.loc[11, constants.Machines.xsubtr],
			constants.G74.x11-constants.G74.tx_x,
			places=4
		)

	def test_adding_machines_psse(self):
		"""
			Confirm that machine parameters correctly added to DataFrame
		"""
		# Get initial machine data
		machine_data = test_module.MachineData()
		machine_data.update()
		machine_data_initial = machine_data.df

		# Add machines to model
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()
		g74_data.add_machines()

		# Get DataFrame of updated machines
		machine_data.update()
		machine_data_final = machine_data.df

		# Check that all new machines have been added correctly
		self.assertEqual(
			len(machine_data_final),
			len(machine_data_initial)+len(g74_data.df_machines)
		)

	def test_calculate_machine_time_dependant_ac_contribution_iec(self):
		"""
			Function tests adding / updating machines for different time steps
			and determines whether the fault contribution from a particular machine
			matches with the expected values.  Machine fault contribution is
			determined using the IEC method (pssarrays.iecs_currents)

			Calculation is performed for AC component only
		:return:
		"""
		# Reload SAV case
		self.psse.load_data_case()

		# Bus numbers to test - These are busbars with motors directly contributing to a fault at
		# this busbar
		buses_to_test = [11, 33]
		fault_times_to_test = (0.0001, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.1, 0.11, 0.12)
		# Pre fault voltage = 1.0p.u. because modelled as a slack busbar in PSSe test model
		pre_fault_v = 1.0
		# Load value modelled as 5.0 MVA in PSSe test model
		load_value = 5.0

		# Model setup
		# Add machines to model
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		# IEC method for fault current calculations
		iec = test_module.IecFaults(psse=self.psse, buses=buses_to_test)

		# Iterative loop that tests a range of fault times
		dfs = list()
		expected_f33_infeed_pu = list()
		expected_f11_infeed_pu = list()

		# R values remain constant for all fault times
		new_r33_value = constants.G74.rpos
		new_r11_value = new_r33_value - constants.G74.tx_r
		for fault_time in fault_times_to_test:
			# Calculate the expected X value and expected fault in feed in per unit
			# Target busbar is 11kV so need to account for 11/33kV transformer impedance
			if fault_time > constants.PSSE.min_fault_time:
				new_x33_value = 1.0 / ((1.0 / constants.G74.x11) * math.exp(-fault_time / constants.G74.t11))
			else:
				new_x33_value = constants.G74.x11

			new_x11_value = new_x33_value - constants.G74.tx_x

			expected_f33_infeed_pu.append(pre_fault_v/(new_r33_value**2 + new_x33_value**2)**0.5)
			expected_f11_infeed_pu.append(pre_fault_v/(new_r11_value**2 + new_x11_value**2)**0.5)
			# Calculate new machine impedance values
			g74_data.calculate_machine_impedance(fault_time=fault_time)

			# Confirm value is correct for two difference values
			self.assertAlmostEqual(
				g74_data.df_machines.loc[11, constants.Machines.xsubtr], new_x11_value, places=5)
			self.assertAlmostEqual(
				g74_data.df_machines.loc[11, constants.Machines.xsynch], new_x11_value, places=5)
			self.assertAlmostEqual(
				g74_data.df_machines.loc[33, constants.Machines.xsubtr], new_x33_value, places=5)
			self.assertAlmostEqual(
				g74_data.df_machines.loc[33, constants.Machines.xsynch], new_x33_value, places=5)

			# Add machines to model
			g74_data.add_machines()

			if fault_time == constants.PSSE.min_fault_time:
				self.psse.save_data_case(pth_sav=SAV_CASE_COMPLETE2)

			# Run fault study for both busbars at this time step
			df = iec.fault_study(fault_time=fault_time)
			dfs.append(df)

		# Combine DataFrames into an overall list
		df_all = pd.concat(dfs, axis=1, keys=fault_times_to_test)

		# Calculate the values that would be expected based on the fault times tested
		expected_values_11 = dict(zip(
			fault_times_to_test,
			[x*load_value*constants.G74.mva_11/(math.sqrt(3)*11.0) for x in expected_f11_infeed_pu]
		))
		expected_values_33 = dict(zip(
			fault_times_to_test,
			[x*load_value*constants.G74.mva_33/(math.sqrt(3)*33.0) for x in expected_f33_infeed_pu]
		))

		# Extract the RMS symmetrical values and confirm they vary correctly with fault time
		# The ik'' values are used rather than ib since the IEC method takes into consideration other factors in
		# determining ib
		df_ik = df_all.xs(constants.BkdyFileOutput.ik11, axis=1, level=1, drop_level=True)
		cols = df_ik.columns
		# Validate that all values are correct
		for fault_time in cols:
			self.assertAlmostEqual(expected_values_11[fault_time], df_ik.loc[11, fault_time], places=3)
			self.assertAlmostEqual(expected_values_33[fault_time], df_ik.loc[33, fault_time], places=3)

	def test_calculate_machine_time_dependant_dc_contribution_iec(self):
		"""
			Function tests adding / updating machines for different time steps
			and determines whether the fault contribution from a particular machine
			matches with the expected values.  Machine fault contribution is
			determined using the IEC method (pssarrays.iecs_currents)

			Validation is performed for DC component which doesn't involve changing the
			motor contribution.
		:return:
		"""
		# Bus numbers to test - These are busbars with motors directly contributing to a fault at
		# this busbar
		buses_to_test = [11, 33]
		fault_times_to_test = (0.0001, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.1, 0.11, 0.12)
		# Pre fault voltage = 1.0p.u. because modelled as a slack busbar in PSSe test model
		pre_fault_v = 1.0
		# Load value modelled as 5.0 MVA in PSSe test model
		load_value = 5.0

		# Calculated ik'' values for 11 and 33kV busbars
		r33_value = constants.G74.rpos
		r11_value = r33_value - constants.G74.tx_r
		x33_value = constants.G74.x11
		x11_value = x33_value - constants.G74.tx_x
		expected_ik33_pu = pre_fault_v / (r33_value**2 + x33_value**2)**0.5
		expected_ik11_pu = pre_fault_v / (
				r11_value**2 +
				x11_value**2)**0.5

		# Empty dictionaries populated with each time step
		expected_values_11 = dict()
		expected_values_33 = dict()

		# Model setup
		# Add machines to model
		g74_data = test_module.G74FaultInfeed()
		g74_data.identify_machine_parameters()
		g74_data.calculate_machine_mva_values()

		# For DC component machine parameters do not change with time since based on Ik''
		g74_data.calculate_machine_impedance(fault_time=0.0)
		# Add machines to model now since parameters do not change
		g74_data.add_machines()

		# IEC method for fault current calculations
		iec = test_module.IecFaults(psse=self.psse, buses=buses_to_test)

		# Iterative loop that tests a range of fault times
		dfs = list()
		for fault_time in fault_times_to_test:
			# Calculate the expected DC value at this time using equation 5.3.1 of G74
			expected_values_11[fault_time] = (
					math.sqrt(2) *
					(expected_ik11_pu * load_value * constants.G74.mva_11 / (math.sqrt(3) * 11.0)) *
					math.exp(-2 * math.pi * 50.0 * fault_time * (r11_value / x11_value))
			)
			expected_values_33[fault_time] = (
					math.sqrt(2) *
					(expected_ik33_pu * load_value * constants.G74.mva_33 / (math.sqrt(3) * 33.0)) *
					math.exp(-2 * math.pi * 50.0 * fault_time * (r33_value / x33_value))
			)

			# Run fault study for both busbars at this time step
			df = iec.fault_study(fault_time=fault_time)
			dfs.append(df)

		# Combine DataFrames into an overall list
		df_all = pd.concat(dfs, axis=1, keys=fault_times_to_test)

		# Extract the DC component and confirm vary correctly with fault time
		df_idc = df_all.xs(constants.BkdyFileOutput.idc, axis=1, level=1, drop_level=True)
		cols = df_idc.columns
		# Validate that all values are correct
		for fault_time in cols:
			self.assertAlmostEqual(expected_values_11[fault_time], df_idc.loc[11, fault_time], places=3)
			self.assertAlmostEqual(expected_values_33[fault_time], df_idc.loc[33, fault_time], places=3)

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
