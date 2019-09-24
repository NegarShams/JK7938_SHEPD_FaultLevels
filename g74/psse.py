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


class InductionData:
	def __init__(self, flag=1, sid=-1):
		"""

		:param int flag: (optional=1) - Only in-service machines at in-service busbars
		:param int sid: (optional=-1) - Allows customer region to be defined
		"""
		# DataFrames populated with type and voltages for each study
		# Index of DataFrame is busbar number as an integer
		self.df = pd.DataFrame()

		# constants
		self.logger = logging.getLogger(constants.Logging.logger_name)

		self.flag = flag
		self.sid = sid

		self.count = -1

	def add_to_idev(self, target):
		"""
			Function will add impendance data for machines if none exist
		:param str target: Existing idev file to append machine impedance data to and close
		:return None:
		"""

		if self.get_count() > 0:
			self.logger.error('There are induction machines in the model but the script has not been developed to '
							  'take these into account')

		# Append extra 0 to end of line
		with open(target, 'a') as csv_file:
			csv_file.write('0')

		return None

	def get_count(self, reset=False):
		"""
			Updates induction machine data from SAV case
		"""
		# Only checks if not already empty
		if self.count == -1 or reset:
			# Declare functions
			func_count = psspy.aindmaccount

			# Retrieve data from PSSE
			ierr_count, number = func_count(
				sid=self.sid,
				flag=self.flag)

			if ierr_count > 0:
				self.logger.critical(
					(
						'Unable to retrieve the number of induction machines in the PSSE SAV case and the following error '
						'codes {} from the functions <{}>'
					).format(ierr_count, func_count.__name__)
				)
				raise SyntaxError('Error importing data from PSSE SAV case')

			self.count = number
		return self.count


class MachineData:
	"""
		Class will contain all of the Machine Data
	"""
	def __init__(self, flag=2, sid=-1):
		"""

		:param int flag:
		:param int sid:
		"""
		self.sid = sid
		self.flag = flag
		self.logger = logging.getLogger(constants.Logging.logger_name)

		self.c = constants.Machines

		self.df = pd.DataFrame()

	def update(self):
		"""
			Update dataframe with the data necessary for the idev file
		:return None:
		"""
		# Declare functions
		func_int = psspy.amachint
		func_real = psspy.amachreal
		func_char = psspy.amachchar

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus, ))
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.x_subtr, self.c.x_trans, self.c.x_synch))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.identifier,))

		if ierr_int > 0 or ierr_char > 0 or ierr_real > 0:
			self.logger.critical(
				(
					'Unable to retrieve the busbar type codes from the SAV case and PSSE returned the '
					'following error codes {}, {} and {} from the functions <{}>, <{}> and <{}>'
				).format(ierr_int, ierr_real, ierr_char, func_int.__name__, func_real.__name__, func_char.__name__)
			)
			raise SyntaxError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		data = iarray + rarray + carray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.bus, self.c.x_subtr, self.c.x_trans, self.c.x_synch, self.c.identifier]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		df = pd.DataFrame(data).transpose()
		df.columns = initial_columns

		self.df = df

		return None

	def produce_idev(self, target):
		"""
			Produces an idev file based on the data in the PSSE model in the appropriate format for importing into
			the BKDY fault current calculation method
		:param str target:  Target path to save the idev file to as a csv
		:return None:
		"""

		self.update()

		df = self.df

		df[self.c.x11] = df[self.c.x_subtr]
		df[self.c.x1d] = df[self.c.x_trans]
		df[self.c.xd] = df[self.c.x_synch]

		self.logger.debug(
			(
				'Default time constant values assumed for all machines and q axis reactance value '
				'all assumed to be equal to d axis reactance values.\n'
				'{} = {}, {} = {}\n'
				'{} = {}, {} = {}'
			).format(
				self.c.t1d0, constants.SHEPD.t1d0, self.c.t1q0, constants.SHEPD.t1q0,
				self.c.t11d0, constants.SHEPD.t11d0, self.c.t11q0, constants.SHEPD.t11q0
			)
		)
		df[self.c.t1d0] = constants.SHEPD.t1d0
		df[self.c.t11d0] = constants.SHEPD.t11d0
		df[self.c.t1q0] = constants.SHEPD.t1q0
		df[self.c.t11q0] = constants.SHEPD.t11q0
		df[self.c.x1q] = df[self.c.x1d]
		df[self.c.xq] = df[self.c.xd]

		# Reorder columns into format needed for idev file
		df = df[self.c.bkdy_col_order]

		# Export to a cav file
		df.to_csv(target, header=False, index=False)

		# Add in empty 0 to mark the end of the file
		with open(target, 'a') as csv_file:
			csv_file.write('0\n')

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

		# Status flag for whether SAV case is converted or not
		self.converted = False

		# Flag that is set to True if any of the errors that occur could affect the accuracy of the BKDY calculated
		# fault levels
		self.bkdy_issue = False

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

		self.converted = False

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

	def convert_sav_case(self):
		"""
			To make it possible to carry out fault current studies using the BKDY method it is necessary
			to convert the generators and loads to norton equivalent sources
			(see 10.12 of PSSE POM v33)
			This module converts the SAV case and ensures a flag is set determining the status.
			Once converted, SAV case needs to be reloaded to restore to original form
		:return None:
		"""
		if not self.converted:
			self.convert_gen()
			self.convert_load()
			self.converted = True
		else:
			self.logger.debug('Attempted call to convert generation and load but are already converted')

		# Generators will now be ordered
		func_ordr = psspy.ordr
		ierr = func_ordr(opt=0)
		if ierr > 0:
			self.logger.critical(
				(
					'Error ordering the busbars into a sparsity matrix using function <{}> which returned '
					'error code {}'
				).format(func_ordr.__name__, ierr)
			)

		# Factorize admittance matrix
		func_fact = psspy.fact
		ierr = func_fact()
		if ierr > 0:
			self.logger.critical(
				(
					'Error when trying to factorize admittance matrix using function <{}> which returned error code {}'
				).format(func_fact.__name__, ierr)
			)

		# TODO: TYSL - Do not believe this is required

		return None

	def convert_gen(self):
		"""
			Script to control the conversion of generation
		:return None: Will only get to end if successful
		"""
		self.logger.debug('Converting generation in model ready for BKDY study')
		# Defines the option for psspy.cong with regards to treatment of conventional machines and induction machines
		# 0 = Uses Zsorce for conventional machines
		# 1 = Uses X'' for conventional machines
		# 2 = Uses X' for conventional machines
		# 3 = Uses X for conventional machines
		x_type = 0

		# Check that no induction machines exist since otherwise assumptions above are not applicable
		if InductionData().get_count() > 0:
			self.logger.warning(
				(
					'There are induction machines included in the PSSE sav case {}.  Some of the assumptions in the '
					'these scripts may no longer be valid.'
				).format(self.sav_name)
			)
			self.bkdy_issue = True

		# Convert generators to suitable equivalent ready for study, only if not already converted
		func_cong = psspy.cong
		ierr = func_cong(opt=x_type)

		if ierr == 1 or ierr == 5:
			self.logger.critical(
				'Critical error in execution of function <{}> which returned the error code {}'
			).format(func_cong.__name__, ierr)
			raise SyntaxError('Scripting error when trying to convert generators')
		elif ierr == 2:
			self.logger.warning(
				'Attempted to convert generators when already converted, this is not an issue but indicates a script '
				'issue.'
			)
		elif ierr == 3 or ierr == 4:
			self.logger.error(
				(
					'Conversion of generators occurred due to incorrect machine impedances or stalled induction '
					'machines.  The function <{}> returned error code {} and you are suggested to check the '
					'contents of the SAV case {}'
				).format(func_cong.__name__, ierr, self.sav)
			)
			raise ValueError('Unable to convert generators')

		return None

	def convert_load(self):
		"""
			Script to control the conversion of load
		:return None:  Will only get to end if successful
		"""
		self.logger.debug('Converting loads in model ready for BKDY study')

		# Method of conversion of loads
		status1 = 0  # If set to 1 or 2 then loads are reconstructed
		# Whether loads connected to some busbars should be skipped
		status2 = 0  # If set to 1 then only type 1 buses, if set to 2 then type 2 and 3 buses

		# TODO: Sensitivty check to determine if these need to be available as an input
		# Constants used to define the way that loads are treated in the conversion
		# Loads converted to constant admittance in active and reactive power
		loadin1 = 0.0
		loadin2 = 1.0
		loadin3 = 0.0
		loadin4 = 1.0

		func_conl = psspy.conl
		# Multiple runs of the function are necessary to convert the loads
		# Run 1 = Initialise for load conversion
		run_count = 1
		ierr, _ = func_conl(
			sid=self.sid,
			apiopt=run_count,
			status1=status1
		)
		if ierr > 0:
			self.logger.critical(
				(
					'Unable to convert the loads using function <{}> which returned error code {} for sav case {} '
					'during load conversion {}'
				).format(func_conl.__name__, ierr, self.sav, run_count)
			)
			raise ValueError('Unable to convert loads')

		# Run 3 = Post processing house keeping
		run_count = 2

		# Ensures that a second run is carried out if unconverted loaded remain in the system model
		unconverted_loads = 1
		i = 0
		while unconverted_loads > 0:
			i += 1
			ierr, unconverted_loads = func_conl(
				sid=self.sid,
				all=1,
				apiopt=run_count,
				status2=status2,
				loadin1=loadin1,
				loadin2=loadin2,
				loadin3=loadin3,
				loadin4=loadin4
			)
			if ierr > 0:
				self.logger.critical(
					(
						'Unable to convert the loads using function <{}> which returned error code {} for sav case {} '
						'during load conversion {}'
					).format(func_conl.__name__, ierr, self.sav, run_count)
				)
				raise ValueError('Unable to convert loads')

			# Catch in case stuck in infinite loop
			if i > 3:
				self.logger.critical(
					(
						'Trying to convert loads resulted in lots of calls to <{}> with apiopt={}.  In total {} '
						'iterations took place and still the number of unconverted loads == {}'
					).format(func_conl.__name__, run_count, i, unconverted_loads)
				)
				raise SyntaxError('Uncontrolled iteration')

		# Run 2 = Convert the loads
		run_count = 3
		ierr, _ = func_conl(
			sid=self.sid,
			apiopt=run_count
		)
		if ierr > 0:
			self.logger.critical(
				(
					'Unable to convert the loads using function <{}> which returned error code {} for sav case {} '
					'during load conversion {}'
				).format(func_conl.__name__, ierr, self.sav, run_count)
			)
			raise ValueError('Unable to convert loads')

		return None


class BkdyFaultStudy:
	"""
		Class that contains all the routines necessary for the BKDY fault study method
	"""
	def __init__(self, psse_control):
		"""
			Function deals with the processing of all the routines necessary to calculate the fault currents using the BKDY method
		:param PsseControl psse_control:
		"""
		self.psse = psse_control
		# Subsystem used for selecting all the busbars
		self.sid = 1

		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.breaker_duty_file = str()

	def create_breaker_duty_file(self, target_path):
		self.breaker_duty_file = target_path

		mac_data = MachineData()
		mac_data.produce_idev(target=target_path)
		induction_machines = InductionData()
		induction_machines.add_to_idev(target=target_path)

	def main(self):
		"""
			Main calculation processes
		:return:
		"""
		# TODO: Define bus subsystem to only return faults for particular buses
		# TODO:

		# Convert model
		self.psse.convert_sav_case()

		func_bkdy = psspy.bkdy
		# TODO: ALL needs to be defined with an input once the bus subsystem has been defined
		# TODO Fault time to be defined
		ierr = func_bkdy(
			sid=self.sid,
			all=1,
			apiopt=1,
			lvlbak=-1,
			flttim=constants.fault_time,
			bfile=self.breaker_duty_file)

		if ierr > 0:
			self.logger.critical(
				(
					'Error occured trying to calculate BKDY which returned the following error code {} from the function '
					'<{}>'
				).format(ierr, func_bkdy.__name__)
			)
