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

# Project specific imports
import g74.constants as constants

# Generic python package imports
import sys
import os
import logging
import pandas as pd


class InitialisePsspy:
	"""
		Class to deal with the initialising of PSSE by checking the correct directory is being referenced and has been
		added to the system path and then attempts to initialise it
	"""
	def __init__(self, psse_version=constants.PSSE.version):
		"""
			Initialise the paths and checks that import psspy works
		:param int psse_version: (optional=34)
		"""

		self.psse = False

		# Get PSSE path
		self.psse_py_path, self.psse_os_path = constants.PSSE().get_psse_path(psse_version=psse_version)
		# Add to system path if not already there
		if self.psse_py_path not in sys.path:
			sys.path.append(self.psse_py_path)

		if self.psse_os_path not in os.environ['PATH']:
			os.environ['PATH'] += ';{}'.format(self.psse_os_path)

		if self.psse_py_path not in os.environ['PATH']:
			os.environ['PATH'] += ';{}'.format(self.psse_py_path)

		global psspy
		# #global pssarrays
		try:
			import psspy
			psspy = reload(psspy)
			self.psspy = psspy
			# #import pssarrays
			# #self.pssarrays = pssarrays
		except ImportError:
			self.psspy = None
			# #self.pssarrays = None

	def initialise_psse(self):
		"""
			Initialise PSSE
		:return bool self.psse: True / False depending on success of initialising PSSE
		"""
		if self.psse is True:
			pass
		else:
			error_code = self.psspy.psseinit()

			if error_code != 0:
				self.psse = False
				raise RuntimeError('Unable to initialise PSSE, error code {} returned'.format(error_code))
			else:
				self.psse = True
				# Disable screen output based on PSSE constants
				self.change_output(destination=constants.PSSE.output[constants.DEBUG_MODE])

		return self.psse

	def change_output(self, destination):
		"""
			Function disables the reporting output from PSSE
		:param int destination:  Target destination, default is to disable which sets it to 6
		:return None:
		"""
		print('PSSE output set to: {}'.format(destination))

		# Disables all PSSE output
		_ = self.psspy.report_output(islct=destination)
		_ = self.psspy.progress_output(islct=destination)
		_ = self.psspy.alert_output(islct=destination)
		_ = self.psspy.prompt_output(islct=destination)

		return None


class BusData:
	"""
		Stores busbar data
	"""

	def __init__(self, flag=2, sid=-1):
		"""

		:param int flag: (optional=2) - Includes in and out of service busbars
		:param int sid: (optional=-1) - Allows customer region to be defined
		"""
		# DataFrames populated with type and voltages for each study
		# Index of DataFrame is busbar number as an integer
		self.df_state = pd.DataFrame()
		self.df_voltage = pd.DataFrame()
		self.df_limits = pd.DataFrame()

		# Populated with list of contingency names where voltages exceeded
		self.voltages_exceeded_steady = list()
		self.voltages_exceeded_step = list()

		# constants
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.Busbars

		self.flag = flag
		self.sid = sid
		self.update()

	def get_voltages(self):
		"""
			Returns the voltage data from the latest load flow
		:return (list, list), (nominal, voltage):
		"""
		func_real = psspy.abusreal
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.nominal, self.c.voltage))

		if ierr_real > 0:
			self.logger.critical(
				(
					'Unable to retrieve the busbar voltage data from the SAV case and PSSE returned '
					'the following error code {} from the function <{}>').format(ierr_real, func_real.__name__)
			)
			raise SyntaxError('Error importing data from PSSE SAV case')

		return rarray[0], rarray[1]

	def update(self):
		"""
			Updates busbar data from SAV case
		"""
		# Declare functions
		func_int = psspy.abusint
		func_char = psspy.abuschar

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus, self.c.state))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus_name,))

		if ierr_int > 0 or ierr_char > 0:
			self.logger.critical(
				(
					'Unable to retrieve the busbar type codes from the SAV case and PSSE returned the '
					'following error codes {} and {} from the functions <{}> and <{}>'
				).format(ierr_int, ierr_char, func_int.__name__, func_char.__name__)
			)
			raise SyntaxError('Error importing data from PSSE SAV case')

		nominal, voltage = self.get_voltages()

		# Combine data into single list of lists
		data = iarray + [nominal] + [voltage] + carray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.bus, self.c.state, self.c.nominal, self.c.voltage, self.c.bus_name]

		state_columns = [self.c.bus, self.c.bus_name, self.c.nominal, self.c.state]
		voltage_columns = [self.c.bus, self.c.bus_name, self.c.nominal, self.c.voltage]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		df = pd.DataFrame(data).transpose()
		df.columns = initial_columns
		df.index = df[self.c.bus]

		# Since not a contingency populate all columns
		self.df_state = df[state_columns]
		self.df_voltage = df[voltage_columns]

		# Insert steady_state voltage limits
		self.add_voltage_limits(df=self.df_voltage)

	def add_voltage_limits(self, df):
		"""
			Function will insert the upper and lower voltage limits for each busbar into the DataFrame
		:param pd.DataFrame df: DataFrame for which voltage limits should be inserted into (must contain busbar nominal
				voltage under the column [c.nominal]
		"""

		# Get voltage limits from constants which are returned as a dictionary
		v_limits = constants.SHEPD.steady_state_limits

		if self.df_limits.empty:
			# Convert to a DataFrame
			self.df_limits = pd.concat([
				pd.DataFrame(data=v_limits.keys()),
				pd.DataFrame(data=v_limits.values())],
				axis=1)
			c = constants.Busbars
			self.df_limits.columns = [c.nominal_lower, c.nominal_upper, c.lower_limit, c.upper_limit]

		# Add limits into busbars voltage DataFrame
		for i, row in self.df_limits.iterrows():
			# Add lower limit values
			df.loc[
				(df[c.nominal] > row[c.nominal_lower]) &
				(df[c.nominal] <= row[c.nominal_upper]),
				c.lower_limit] = row[c.lower_limit]
			# Add upper limit values
			df.loc[
				(df[c.nominal] > row[c.nominal_lower]) &
				(df[c.nominal] <= row[c.nominal_upper]),
				c.upper_limit] = row[c.upper_limit]
		return None


class PsseControl:
	"""
		Class to obtain and store the PSSE data
	"""
	def __init__(self, areas=list(range(0, 100, 1)), sid=-1):
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.sav = str()
		self.sav_name = str()
		self.sid = sid
		self.areas = areas

	def load_data_case(self, pth_sav=None):
		"""
			Load the study case that PSSE should be working with
		:param str pth_sav:  (optional=None) Full path to SAV case that should be loaded
							if blank then it will reload previous
		:return None:
		"""
		try:
			func = psspy.case
		except NameError:
			self.logger.debug('PSSE has not been initialised when trying to load save case, therefore initialised now')
			success = InitialisePsspy().initialise_psse()
			if success:
				func = psspy.case
			else:
				self.logger.critical('Unable to initialise PSSE')
				raise ImportError('PSSE Initialisation Error')

		# Allows case to be reloaded
		if pth_sav is None:
			pth_sav = self.sav
		else:
			# Store the sav case path and name of the file
			self.sav = pth_sav
			self.sav_name, _ = os.path.splitext(os.path.basename(pth_sav))

		# Load case file
		ierr = func(sfile=pth_sav)
		if ierr > 0:
			self.logger.critical(
				(
					'Unable to load PSSE Saved Case file:  {}.\n'
					'PSSE returned the error code {} from function {}'
				).format(pth_sav, ierr, func.__name__)
			)
			raise ValueError('Unable to Load PSSE Case')

		# Set the PSSE load flow tolerances to ensure all studies done with same parameters
		self.set_load_flow_tolerances()

		return None

	def save_data_case(self, pth_sav=None):
		"""
			Load the study case that PSSE should be working with
		:param str pth_sav:  (optional=None) Full path to SAV case that should be loaded
							if blank then it will reload previous
		:return None:
		"""
		func = psspy.save

		# Allows case to be reloaded
		if pth_sav is None:
			pth_sav = self.sav

		# Load case file
		ierr = func(sfile=pth_sav)
		if ierr > 0:
			self.logger.critical(
				(
					'Unable to save PSSE Saved Case to file:  {}.\n PSSE returned the error code {} from function {}'
				).format(pth_sav, ierr, func.__name__)
			)
			raise ValueError('Unable to Save PSSE Case')

		# ## Set the PSSE load flow tolerances to ensure all studies done with same parameters
		# #self.set_load_flow_tolerances()

		return None

	def set_load_flow_tolerances(self):
		"""
			Function sets the tolerances for when performing Load Flow studies
		:return None:
		"""
		# Function for setting PSSE solution parameters
		func = psspy.solution_parameters_4

		ierr = func(
			intgar2=constants.PSSE.max_iterations,
			realar6=constants.PSSE.mw_mvar_tolerance
		)

		if ierr > 0:
			self.logger.warning(
				(
					'Unable to set the max iteration limit in PSSE to {}, the model may struggle to converge but will '
					'continue anyway.  PSSE returned the error code {} from function <{}>'
				).format(constants.PSSE.max_iterations, ierr, func.__name__)
			)
		return None

	def run_load_flow(self, flat_start=False, lock_taps=False):
		"""
			Function to run a load flow on the psse model for the contingency, if it is not possible will
			report the errors that have occurred
		:param bool flat_start: (optional=False) Whether to carry out a Flat Start calculation
		:param bool lock_taps: (optional=False)
		:return (bool, pd.DataFrame) (convergent, islanded_busbars):
			Returns True / False based on convergent load flow existing
			If islanded busbars then disconnects them and returns details of all the islanded busbars in a DataFrame
		"""
		# Function declarations
		if flat_start:
			# If a flat start has been requested then must use the "Fixed Slope Decoupled Newton-Raphson Power Flow Equations"
			func = psspy.fdns
		else:
			# If flat start has not been requested then use "Newton Raphson Power Flow Calculation"
			func = psspy.fnsl

		if lock_taps:
			tap_changing = 0
		else:
			tap_changing = 1

		# Run loadflow with screen output controlled
		c_psse = constants.PSSE
		ierr = func(
			options1=tap_changing,  # Tap changer stepping enabled
			options2=c_psse.tie_line_flows,  # Don't enable tie line flows
			options3=c_psse.phase_shifting,  # Phase shifting adjustment disabled
			options4=c_psse.dc_tap_adjustment,  # DC tap adjustment disabled
			options5=tap_changing,  # Include switched shunt adjustment
			options6=flat_start,  # Flat start depends on status of <flat_start> input
			options7=c_psse.var_limits,  # Apply VAR limits immediately
			# #options7=99,  # Apply VAR limits automatically
			options8=c_psse.non_divergent)  # Non divergent solution

		# Error checking
		if ierr == 1 or ierr == 5:
			# 1 = invalid OPTIONS value
			# 5 = prerequisite requirements for API are not met
			self.logger.critical('Script error, invalid options value or API prerequisites are not met')
			raise SyntaxError('SCRIPT error, invalid options value or API prerequisites not met')
		elif ierr == 2:
			# generators are converted
			self.logger.critical(
				'Generators have been converted, you must reload SAV case or reverse the conversion of the generators')
			raise IOError('The generators are converted and therefore it is not possible to run loadflow')
		elif ierr == 3 or ierr == 4:
			self.logger.error('Error there are islanded busbars and so a convergent load flow was not possible')
			# TODO:  Can implement TREE to identify islanded busbars but will then need to ensure restored
			# buses in island(s) without a swing bus; use activity TREE
			# #islanded_busbars = self.get_islanded_busbars()
			# #return False, islanded_busbars
		elif ierr > 0:
			# Capture future errors potential from a change in PSSE API
			self.logger.critical('UNKNOWN ERROR')
			raise SyntaxError('UNKNOWN ERROR')

		# Check whether load flow was convergent
		convergent = self.check_convergent_load_flow()

		return convergent, pd.DataFrame()

	def check_convergent_load_flow(self):
		"""
			Function to check if the previous load flow was convergent
		:return bool convergent:  True if convergent and false if not
		"""
		error = psspy.solved()
		if error == 0:
			convergent = True
		elif error in (1, 2, 3, 5):
			self.logger.debug(
				'Non-convergent load flow due to a non-convergent case with error code {}'.format(error)
			)
			convergent = False
		else:
			self.logger.error(
				'Non-convergent load flow due to script error or user input with error code {}'.format(error)
			)
			convergent = False

		return convergent
