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

# Project specific imports
import g74.constants as constants

# Generic python package imports
import sys
import os
import logging
import pandas as pd
import numpy as np
import math
import time
import re

# Version of PSSE that will be initialised
DEFAULT_PSSE_VERSION = 33

# TODO: Report error for busbars which do not actually exist in model


def extract_values(line, expected_length=0):
	"""
		Extract values from line and if can be converted to a float return as list
	:param str line:  Line to be processed and split so that its float values can be extracted
	:param int expected_length: (optional) Used to check that the number of parameters returned matches
								the expected number
	:return list extracted:  List of values that have been extracted
	"""
	logger = logging.getLogger(constants.Logging.logger_name)

	# Check for any NaN values in line and if so convert to 0.0
	# TODO: Add error message reporting to warn user that this has happened
	line = line.replace(constants.BkdyFileOutput.nan_value, constants.BkdyFileOutput.nan_replacement)

	reg_search = constants.BkdyFileOutput.reg_search
	# reg find all returns a list of lists showing which of the search items matched
	data = reg_search.findall(line)
	# Convert into a flat list, removing the blank strings and converting to floats
	# #extracted = [float(item) for sublist in data for item in sublist if item != '']
	# Convert into a flat list, removing the blank strings and converting to floats
	extracted = list()
	# TODO: Looping through multiple loops here, maybe better way to return from REGEX search
	value = np.nan
	for group in data:
		# #TODO: Define as inputs rather than scripts
		if (
				constants.BkdyFileOutput.nan_term1 in group or
				constants.BkdyFileOutput.nan_term2 in group or
				constants.BkdyFileOutput.nan_term3 in group
		):
			value = 0.0
		else:
			# Extract single value and convert to flat list
			for item in group:
				if item != '':
					value = float(item)
		extracted.append(value)

	# Error processing to check if lengths as expected
	if (
			len(extracted) != expected_length and
			expected_length > 0
			and constants.BkdyFileOutput.infinity_error not in line
			and 'Infinity' not in line
	):
		logger.error(
			(
				'Processing of the string below into a list of floats returned a list of values '
				'different to what was expected.\n'
				'Therefore it is not possible to associate the specific list items to parameters\n'
				'{}\n'
				'The following values were extracted:\n'
				'{}'
			).format(line, extracted)
		)
		raise ValueError('Not possible to reliably process results, see log output')

	return extracted


class InitialisePsspy:
	"""
		Class to deal with the initialising of PSSE by checking the correct directory is being referenced and has been
		added to the system path and then attempts to initialise it
	"""
	def __init__(self, psse_version=DEFAULT_PSSE_VERSION):
		"""
			Initialise the paths and checks that import psspy works
		:param int psse_version: (optional=34)
		"""

		self.psse = False
		self.c = constants.PSSE
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.psse_version = psse_version

		# Set PSSE paths for defined PSSE version
		self.psse_py_path = str()
		self.psse_os_path = str()
		self.set_psse_path()

		global psspy
		global redirect
		global pssarrays
		global sliderPy
		try:
			# Import psspy used for manipulating PSSE
			import psspy
			import redirect
			psspy = reload(psspy)
			redirect = reload(redirect)
			self.psspy = psspy

		except ImportError:
			self.psspy = None
			self.logger.error(
				(
					'Unable to initialise PSSPY which is thought to be installed in the directory {}, suggest '
					'checking for {} in this directory'
				).format(self.psse_py_path, self.c.psspy_to_find)
			)
			raise ImportError('Unable to initialise PSSPY and therefore cannot run PSSE studies')

		# TODO: Better error handling / importing to only include this when needed
		try:
			# Import pssarrays used for data extraction from PSSE
			import pssarrays
			pssarrays = reload(pssarrays)
			self.pssarrays = pssarrays
		except ImportError:
			self.pssarrays = None
			self.logger.error(
				(
					'Unable to initialise PSSARRAYS which is used to process IEC results and thought to be '
					'installed in the directory {}, suggest checking for {} in this directory'
				).format(self.psse_py_path, self.c.pssarrays_to_find)
			)
			raise ImportError('Unable to initialise PSSARRAYS used for IEC fault current data extraction')

		# Separate error handling import since need is not essential
		try:
			# Import pssarrays used for data extraction from PSSE
			import sliderPy
			sliderPy = reload(sliderPy)
			self.sliderPy = sliderPy
		except ImportError:
			self.sliderPy = None
			self.logger.warning(
				'Unable to import sliderPy and therefore cannot interact with the PSSE sliders to obtain busbar'
				'numbers.  The script will continue to work but will not be pre-populated with busbars.'
			)

	def set_psse_path(self, reset=False):
		"""
			Function returns the PSSE path specific to this version of psse
		:param bool reset: (optional=False) - If set to True then class is reset with a new psse_version
		:return str self.psse_path:
		"""
		self.logger.debug('Adding PSSE paths to windows environment')
		if self.psse_py_path and self.psse_os_path and not reset:
			return self.psse_py_path, self.psse_os_path

		# Produce directory for standard installation
		self.psse_py_path = os.path.join(self.c.program_files_directory, self.c.psse_paths[self.psse_version])
		self.psse_os_path = os.path.join(self.c.program_files_directory, self.c.os_paths[self.psse_version])
		
		# Check if these paths actually exist and if not then carry out a search for PSSE
		if not os.path.exists(self.psse_py_path) and not os.path.exists(self.psse_os_path):
			t0 = time.time()
			self.logger.info('PSSE not installed in default directories and so searching for installed location')
			self.psse_py_path, self.psse_os_path = self.find_psspy(start_directory=self.c.default_install_directory)
			if not self.psse_py_path or not self.psse_os_path:
				self.logger.error('Unable to find PSSE installation, will attempt to continue but likely to fail')
			self.logger.info('Took {:.2f} seconds to find PSSE'.format(time.time()-t0))

		# Add to system path if not already there
		if self.psse_py_path not in sys.path:
			sys.path.append(self.psse_py_path)

		if self.psse_os_path not in os.environ['PATH']:
			os.environ['PATH'] += ';{}'.format(self.psse_os_path)

		if self.psse_py_path not in os.environ['PATH']:
			os.environ['PATH'] += ';{}'.format(self.psse_py_path)
		
		return self.psse_py_path, self.psse_os_path

	def find_psspy(self, start_directory='C:\\'):
		"""
			Function to search entire directory and find PSSE installation
		:param str start_directory:  Directory from which to start the search
		:return (str, str) (self.psse_py_path, self.psse_os_path):  Returns that paths to PSSE python and executable
		"""
		# Initialise variables
		psse_py_path = str()
		psse_os_path = str()

		# Produce list of drives to search
		# #drives = list()
		# #for letter in string.ascii_uppercase:
		# #	drive = '{}:\\'.format(letter)
		# #	if os.path.isdir(drive):
		# #		drives.append(drive)

		# Executables that relate to PSSE
		# #for drive in drives:
		# #	self.logger.debug('Searching drive {}'.format(drive))
		for root, dirs, files in os.walk(start_directory):  # Walks through all subdirectories searching for file
			[dirs.remove(d) for d in list(dirs) if d.startswith('$') or d.startswith('.')]
			# TODO: Need different way to confirm which version of PSSE is installed, currently just assumes the
			# TODO: relevant version but yet script has not been tested with PSSE v33+
			if self.c.psspy_to_find in files:
				psse_py_path = root
			elif self.c.psse_to_find in files:
				psse_os_path = root

			if psse_py_path and psse_os_path:
				break

		return psse_py_path, psse_os_path
	
	def initialise_psse(self, running_from_psse=False):
		"""
			Initialise PSSE
		:param bool running_from_psse: (optional=False) - If set to True then running in PSSE and so output won't be
									redirected to PSSE
		:return bool self.psse: True / False depending on success of initialising PSSE
		"""
		if self.psse is True:
			pass
		else:
			# Redirect statement ensures that psse output goes to python and avoids popup windows only if running in
			# Python
			if not running_from_psse:
				redirect.psse2py()
			error_code = self.psspy.psseinit()

			if error_code != 0:
				self.psse = False
				raise RuntimeError('Unable to initialise PSSE, error code {} returned'.format(error_code))
			else:
				self.psse = True

				# ## Disable screen output based on PSSE constants
				# #self.change_output(destination=constants.PSSE.output[constants.DEBUG_MODE])

		return self.psse


class BusData:
	"""
		Stores busbar data
	"""

	def __init__(self, flag=1, sid=-1):
		"""
		:param int flag: (optional=1) - Include only in-service busbars
		:param int sid: (optional=-1) - Allows customer region to be defined
		"""
		# DataFrames populated with type and voltages for each study
		# Index of DataFrame is busbar number as an integer
		self.df = pd.DataFrame()

		# Populated with list of contingency names where voltages exceeded
		self.voltages_exceeded_steady = list()
		self.voltages_exceeded_step = list()

		# constants
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.Busbars

		self.flag = flag
		self.sid = sid
		self.update()

	def update(self):
		"""
			Updates busbar data from SAV case
		"""
		# Declare functions
		func_int = psspy.abusint
		func_real = psspy.abusreal
		func_char = psspy.abuschar

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus, self.c.state))
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.nominal, self.c.voltage))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus_name,))

		if ierr_int > 0 or ierr_char > 0 or ierr_real > 0:
			self.logger.critical(
				(
					'Unable to retrieve the busbar data from the SAV case and PSSE returned the '
					'following error codes {}, {} and {} from the functions <{}>, <{}> and <{}>'
				).format(
					ierr_int, ierr_char, ierr_real,
					func_int.__name__, func_char.__name__, func_real.__name__
				)
			)
			raise SyntaxError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		data = iarray + rarray + carray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.bus, self.c.state, self.c.nominal, self.c.voltage, self.c.bus_name]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		df = pd.DataFrame(data).transpose()
		df.columns = initial_columns
		df.index = df[self.c.bus]

		# Since not a contingency populate all columns
		self.df = df


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
			Function will add impedance data for machines if none exist
		:param str target: Existing idev file to append machine impedance data to and close
		:return None:
		"""

		if self.get_count() > 0:
			self.logger.error(
				'There are induction machines in the model but the script has not been developed to '
				'take these into account'
			)
			raise SyntaxError('Incomplete script for induction machines')

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


class PlantData:
	"""
		Class will contain all of the Machine Data
	"""
	def __init__(self, flag=1, sid=-1):
		"""
		:param int flag: (optional=2) - Returns all in-service plant buses including those with no in-service machines
		:param int sid:
		"""
		self.sid = sid
		self.flag = flag
		self.logger = logging.getLogger(constants.Logging.logger_name)

		self.c = constants.Plant

		self.df = pd.DataFrame()
		self.update()

	def update(self):
		"""
			Update DataFrame with the data necessary for the idev file
		:return None:
		"""
		# Declare functions
		func_int = psspy.agenbusint

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus, self.c.status))

		if sum([ierr_int]) > 0:
			self.logger.critical(
				(
					'Unable to retrieve the plant data from the SAV case and PSSE returned the '
					'following error code {} from the function <{}>'
				).format(ierr_int, func_int.__name__)
			)
			raise SyntaxError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		data = iarray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [self.c.bus, self.c.status]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		df = pd.DataFrame(data).transpose()
		df.columns = initial_columns

		self.df = df

		return None


class MachineData:
	"""
		Class will contain all of the Machine Data
	"""
	def __init__(self, flag=2, sid=-1):
		"""
		:param int flag: (optional=2) - Returns all in service
		:param int sid:
		"""
		self.sid = sid
		self.flag = flag
		self.logger = logging.getLogger(constants.Logging.logger_name)

		self.c = constants.Machines

		self.df = pd.DataFrame()

	def update(self):
		"""
			Update DataFrame with the data necessary for the idev file
		:return None:
		"""
		# Declare functions
		func_int = psspy.amachint
		func_real = psspy.amachreal
		func_cplx = psspy.amachcplx
		func_char = psspy.amachchar

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus, ))
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.rpos, self.c.xsubtr, self.c.xtrans, self.c.xsynch))
		ierr_cplx, xarray = func_cplx(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.zsource,))
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
		data = iarray + rarray + xarray + carray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		# in case needed
		initial_columns = [
			self.c.bus, self.c.rpos, self.c.xsubtr, self.c.xtrans, self.c.xsynch, self.c.zsource, self.c.identifier
		]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		df = pd.DataFrame(data).transpose()
		df.columns = initial_columns

		# Split out Z source into R source and X source
		df[self.c.rsource] = df[self.c.zsource].real
		df[self.c.xsource] = df[self.c.zsource].imag

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

		df[self.c.x11] = df[self.c.xsubtr]
		df[self.c.x1d] = df[self.c.xtrans]
		df[self.c.xd] = df[self.c.xsynch]

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
		df[self.c.t11d0] = constants.SHEPD.t11q0
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

	def check_machine_data(self):
		"""
			Script runs through all the machines and checks that they all have reasonable R and X data.
			All those missing R data have it added and a warning message is reported to the user
		:return None:
		"""
		func_seq_mac = psspy.seq_machine_data_3
		self.update()

		df_missing_rpos = self.df[self.df[self.c.rpos] <= self.c.min_r_pos]
		# Iterate over each machine and add missing data
		for idx, machine in df_missing_rpos.iterrows():
			rpos = machine[self.c.xsubtr] / self.c.assumed_x_r
			bus = machine[self.c.bus]
			identifier = machine[self.c.identifier]
			ierr = func_seq_mac(
				i=bus,
				id=identifier,
				realar1=rpos
			)
			if ierr > 0:
				self.logger.error(
					(
						'Unable to change the positive sequence resistance value for the machine connected at '
						'busbar <{}> with ID: {} to a value of {:.5f}.  Therefore the overall results may not be '
						'reliable'
					).format(bus, identifier, rpos)
				)
			else:
				self.logger.warning(
					(
						'Machine connected at busbar <{}> with ID: {} has a positive sequence impedance of <= {} '
						'and has therefore been set to {:.5f} which assumes as X/R of {}'
					).format(bus, identifier, self.c.min_r_pos, rpos, self.c.assumed_x_r)
				)

		return None

	def set_rsource_xsource(self, x_type=constants.Machines.xsubtr):
		"""
			Function will loop through and set the R source value == rpos and the X source value == input parameter
		:param str x_type:  (optional) - Source of data to populate Xsource with
		:return None:
		"""
		# Function for changing machine values
		func_mac_data_change = psspy.machine_chng_2
		# Make sure have latest machine values
		self.update()

		# Obtain DataFrame of machines that are Z source impedance data
		df_missing_zsorce = self.df[
			(self.df[self.c.rsource] != self.df[self.c.rpos]) |
			(self.df[self.c.xsource] != self.df[x_type])
		]

		# Loop through each machine and add missing data
		for idx, machine in df_missing_zsorce.iterrows():
			rsource = machine[self.c.rpos]
			xsource = machine[x_type]
			bus = machine[self.c.bus]
			identifier = machine[self.c.identifier]
			ierr = func_mac_data_change(
				i=bus,
				id=identifier,
				realar8=rsource,
				realar9=xsource
			)
			if ierr > 0:
				self.logger.error(
					(
						'Unable to change the R or X source values for the machine connected at '
						'busbar <{}> with ID: {} to a values of {:.5f} and {:.5f}.  Therefore the overall results '
						'may not be reliable'
					).format(bus, identifier, rsource, xsource)
				)
			else:
				self.logger.info(
					(
						'Machine connected at busbar <{}> with ID: {} has had R and X source values changed to '
						'{:.5f} and {:.5f} based on the values used for {} and {}.'
					).format(bus, identifier, rsource, xsource, self.c.rpos, x_type)
				)

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

		# Boolean value used to determine whether output should be set to PSSE or Python
		self.run_in_psse = None
		# Determine whether running from PSSE or not
		self.running_from_psse()

	def change_output(self, destination=constants.PSSE.output_default):
		"""
			Function disables the reporting output from PSSE
		:param int destination:  (optional=1) Target destination, default is to restore it to 1
		:return None:
		"""
		self.logger.debug('PSSE general output changed to destination = {} (1=default, 6=none)'.format(destination))
		# Disables all PSSE output
		_ = psspy.report_output(islct=destination)
		_ = psspy.progress_output(islct=destination)
		_ = psspy.alert_output(islct=destination)
		_ = psspy.prompt_output(islct=destination)

		print('PSSE output set to: {} and progress output set to: {}'.format(destination, destination))

		return None

	def toggle_progress_output(self, destination):
		"""
			Sets progress output to the destination folder unless running in DEBUG mode
		:param int destination:
		:return None:
		"""
		self.logger.debug('Process output changes to {}'.format(destination))
		progress_destination = min(destination, constants.PSSE.output[constants.DEBUG_MODE])
		_ = psspy.progress_output(islct=progress_destination)

	def running_from_psse(self):
		"""
			Determine if running from PSSE or Python
		:return None:
		"""
		if self.run_in_psse is None:
			# Determine if this script is being run from PSSE or plain Python.
			full_path_executable = sys.executable
			# Remove the folder path and keep only the executable file (in lower case).
			executable = os.path.basename(full_path_executable).lower()

			self.run_in_psse = True
			if executable in ['python.exe', 'pythonw.exe']:
				# If the executable was one of the above, it is a Python session.
				self.run_in_psse = False

	def get_current_sav_case(self):
		"""
			Retrieves the full path to the active sav case if one exists
		:return str sav_case:
		"""
		sav_case, _ = psspy.sfiles()
		self.logger.debug('PSSE currently active save case is: {}'.format(sav_case))

		return sav_case

	def load_data_case(self, pth_sav=None):
		"""
			Load the study case that PSSE should be working with
		:param str pth_sav:  (optional=None) Full path to SAV case that should be loaded
							if blank then it will reload previous
		:return None:
		"""
		# Determine whether being run from PSSE or being run from Python
		self.running_from_psse()

		try:
			func = psspy.case
		except NameError:
			self.logger.debug('PSSE has not been initialised when trying to load save case, therefore initialised now')
			success = InitialisePsspy().initialise_psse(running_from_psse=self.run_in_psse)
			if success:
				func = psspy.case
			else:
				self.logger.critical('Unable to initialise PSSE')
				raise ImportError('PSSE Initialisation Error')

		# Set PSSE output accordingly
		self.change_output(destination=constants.PSSE.output[constants.DEBUG_MODE])

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

		# Set parameters for output values
		self.set_outputs()

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

	def set_outputs(self):
		"""
			Function sets the output parameters to consistent values for subsequent processing
			Options are set based on the values defined in constants
		:return None:
		"""

		func_sc_units = psspy.short_circuit_units
		ierr_sc_units = func_sc_units(ival=constants.PSSE.def_short_circuit_units)

		func_sc_coordinates = psspy.short_circuit_coordinates
		ierr_sc_coordinates = func_sc_coordinates(ival=constants.PSSE.def_short_circuit_coordinates)

		# Set the maximum number of lines for the printing output
		func_lines = psspy.lines_per_page_one_device
		ierr_lines = func_lines(1, 100000)

		if sum([ierr_sc_units, ierr_sc_coordinates, ierr_lines]) > 0:
			self.logger.critical(
				(
					'Unable to change short circuit units to physical (ival={}) with function <{}> returned the '
					'error code {} or short circuit coordinates to polar (ival={}) with function <{}> which returned '
					'the error code {}.  The function <{}> to increase the number of lines on the page '
					'returned the error code {}'
				).format(
					constants.PSSE.def_short_circuit_units, func_sc_units.__name__, ierr_sc_units,
					constants.PSSE.def_short_circuit_coordinates, func_sc_coordinates.__name__,
					ierr_sc_coordinates,
					func_lines.__name__, ierr_lines
				)
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

		else:
			self.logger.debug('Attempted call to convert generation and load but are already converted')

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
		x_type = constants.Machines.bkdy_machine_type

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

		# TODO: Sensitivity check to determine if these need to be available as an input
		# Constants used to define the way that loads are treated in the conversion
		# Loads converted to constant admittance in active and reactive power
		loadin1 = 0.0
		loadin2 = 100.0
		loadin3 = 0.0
		loadin4 = 100.0

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

	def define_bus_subsystem(self, buses, sid=constants.PSSE.sid):
		"""
			Function to define the bus subsystem that will then be used for returning fault current data
		:param int sid: (optional=1)
		:param list buses: List of busbar numbers to be added to the subsystem for fault analysis
		:return int sid:  Subsystem identified number
		"""
		# PSSE functions
		func_subsys_init = psspy.bsysinit
		func_subsys_add = psspy.bsyso

		num_buses = len(buses)
		# Check number of busbars is enough otherwise just define as entire subsystem
		if num_buses == 0:
			self.sid = -1
			self.logger.warning(
				(
					'No busbars provided as an input and therefore no bus subsystem to define,'
					'sid = {}'
				).format(self.sid)
			)
			return self.sid

		# Initialise desired bus subsystem
		ierr_init = func_subsys_init(sid=sid)
		if ierr_init == 1:
			msg0 = (
				(
					'Attempted to use PSSE sid of {} which is outside of the allowable limits of {} to {} and so the '
					'function <{}> returned an error code of {}.'
				).format(sid, 0, 11, func_subsys_init.__name__, ierr_init)
			)

			# Try with a different sid value
			sid = constants.PSSE.sid
			ierr = func_subsys_init(sid=sid)
			if ierr == 0:
				msg1 = 'Instead an sid value of {} has been used'.format(sid)
				self.logger.warning('{}\n{}'.format(msg0, msg1))
			else:
				msg1 = (
					'However even using an sid value of {} as defined in <constants.PSSE> has not resolved the '
					'issue'
				).format(sid)
				self.logger.critical('{}\n{}'.format(msg0, msg1))
				raise ValueError('Not possible to define subsystem')

		# Loop through each bus and add to bus subsystem
		# Seems to produce an error if done via the sub-system definition method
		for bus in buses:
			ierr = func_subsys_add(sid=sid, busnum=bus)
			if ierr == 0:
				self.logger.debug('Busbar {} added to bus subsystem with SID = {}'.format(bus, sid))
			else:
				self.logger.critical(
					(
						'Unable to add busbar {} to subsystem with SID = {} and function '
						'<{}> returned the following error code {}'
					).format(bus, sid, func_subsys_add.__name__, ierr)
				)
				raise ValueError('Unable to add busbar to subsystem')

		self.sid = sid

		return sid

	def write_data_to_psse_report(self, df):
		"""
			Function will export the data in a tabulated for to the PSSE window in exactly the same structure as the
			DataFrame
		:param pd.DataFrame df:
		:return None:
		"""
		# Change report output to the GUI
		# TODO: This script has not been completed
		self.logger.debug('Writing results to PSSE report window')
		if df.empty:
			raise ValueError('DataFrame is empty, nothing to write')
		_ = psspy.report_output(islct=1)


class PsseSlider:
	"""
		Script to deal with interfacing with the PSSE slider diagram to identify attributes, selected items, etc.
	"""
	def __init__(self):
		"""
			Initialise class
		"""
		self.logger = logging.getLogger(constants.Logging.logger_name)

	def get_selected_busbars(self):
		"""
			Function will return a list of all the busbars which have been selected in the slider as integers
		:return list busbars:
		"""
		# Get handle for all components in slider
		doc = sliderPy.GetActiveDocument()
		diagram = doc.GetDiagram()
		components = diagram.GetComponents()

		# TODO: Find name of diagram and number of components to add to log report
		busbars = list()
		for item in components:
			# TODO: Determine if type is busbar and then get name and append to list
			if item.IsSelected() and item.GetComponentType() is sliderPy.ComponentType.Symbol:
				item_details = item.GetMapString()
				# Get busbar number if a busbar rather than circuit
				if 'BU' in item_details:
					try:
						busbars.append(int(re.sub('BU', '', item_details)))
					except ValueError:
						self.logger.warning(
							(
								'Unable to obtain the busbar number from {} which relates to one of the busbars '
								'already selected.  This number will need to be added again manually'
							).format(item_details)
						)

		# Busbars already selected in slider info
		if busbars:
			msg = '\n'.join(['\t - Busbar: {}'.format(bus) for bus in busbars])
			self.logger.info('The following busbars will be added to the fault list:\n{}'.format(msg))
		else:
			self.logger.info('No busbars selected in slider')

		return busbars


class BkdyFaultStudy:
	"""
		Class that contains all the routines necessary for the BKDY fault study method
	"""
	def __init__(self, psse_control):
		"""
			Function deals with the processing of all the routines necessary to calculate the fault currents using
			the BKDY method
		:param PsseControl psse_control:  Handle to PSSE for running of studies
		"""
		self.psse = psse_control
		# Subsystem used for selecting all the busbars
		self.sid = 1
		self.all_buses = 1

		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.breaker_duty_file = str()
		# Dictionary created to relate output names to files
		self.bkdy_files = dict()
		# DataFrame with the combined results for the BKDY method
		self.df_combined_results = pd.DataFrame()

		# List of busbars where there has been an issue in the fault study that are unreliable
		self.unreliable_faulted_buses = list()

		# Check that the MVA values match with the expected value used in the constants
		self.check_mva_value()

	def check_mva_value(self):
		"""
			Function retrieves the MVA value from PSSe and then checks that the value matches
			with the expected value from the constants
			If it does not then displays an error message
		:return None:
		"""
		sysmva = psspy.sysmva()
		if sysmva != constants.BkdyFileOutput.base_mva:
			self.logger.error(
				(
					'The base MVA value of the selected SAV case is {:.1f} but results are expected to '
					'be expressed on {:.1f} MVA base'
				).format(sysmva, constants.BkdyFileOutput.base_mva)
			)

		return None

	def create_breaker_duty_file(self, target_path):
		"""
			Creates the create breaker duty files
		:param str target_path: Target path to save the file to
		:return None:
		"""
		self.breaker_duty_file = target_path

		# Check machine data
		mac_data = MachineData()
		# Check machine data has suitable positive sequence impedance data
		mac_data.check_machine_data()
		if constants.Machines.bkdy_machine_type == 0:
			# Only need to set rsource and xsource values if machine type == 0
			mac_data.set_rsource_xsource()
		mac_data.produce_idev(target=target_path)

		induction_machines = InductionData()
		induction_machines.add_to_idev(target=target_path)

		return None

	def change_report_output(self, destination, output_file=str()):
		"""
			Function disables the reporting output from PSSE
		:param int destination:  Target destination, default is to disable which sets it to 6
		:param str output_file:  Target file to save the output to
		:return None:
		"""

		# Sets PSSE report output to a file rather than general
		func = psspy.report_output
		ierr_report = psspy.report_output(islct=destination, filarg=output_file, options1=0)

		if ierr_report > 0:
			self.logger.critical(
				(
					'Unable to change the report output of psse to the destination: {} using the function <{}> '
					'with the parameters islct={}.  The function returned the following error code: {}'
				).format(output_file, func.__name__, destination, ierr_report)
			)

		return None

	def main(self, name, output_file, fault_time):
		"""
			Main calculation processes
		:param str name: Name to give this result, when combining results this will be used to determine which results
						to extract based on the data included in constants.SHEPD.result
		:param str output_file:  File to store bkdy output into
		:param float fault_time:  Time to use for beaker contact separation
		:return:
		"""
		# TODO: Define bus subsystem to only return faults for particular buses

		# Convert model
		self.psse.convert_sav_case()
		# Function for carrying out the study
		func_bkdy = psspy.bkdy

		# Change destination to file type object
		self.change_report_output(destination=constants.PSSE.output_file, output_file=output_file)

		# Carry out fault current calculation
		ierr = func_bkdy(
			sid=self.sid,
			all=self.all_buses,
			apiopt=1,
			lvlbak=-1,
			flttim=fault_time,
			bfile=self.breaker_duty_file)

		# Change destination back
		# TODO: Could move this to a different component to improve efficiency
		self.change_report_output(destination=constants.PSSE.output[constants.DEBUG_MODE])

		if ierr > 0:
			self.logger.critical(
				(
					'Error occurred trying to calculate BKDY which returned the following error code {} from the function '
					'<{}>'
				).format(ierr, func_bkdy.__name__)
			)

		# Associate this file with the BkdyFile class
		self.bkdy_files[name] = BkdyFile(output_file=output_file, fault_time=fault_time)

	def combine_bkdy_output(self, delete=True):
		"""
			Combines output from bkdy files.
			The particular results that are exported are based on the values detailed in constants.SHEPD.results which
			relate to the name of each result file.  If they cannot be found then all results are exported with the
			particular name appended to each of the headings.
		:param bool delete: (optional=True) - Will delete the original bkdy output files
		:return pd.DataFrame() self.df_combined_results:  DataFrame of the combined results ready for excel export
		"""
		# Empty dictionary that will be populated with DataFrames as they are processed
		dfs = dict()
		# Loops through each of the results and processes the files
		for fault_time, bkdy_file in self.bkdy_files.iteritems():
			self.logger.debug(
				'Processing the BKDY results for fault named: {} and stored in: {}'.format(fault_time, bkdy_file)
			)
			# Extract all data from file and delete file since no longer needed
			df = bkdy_file.process_bkdy_output(delete=delete)
			# #name = '{} {}'.format(fault_time, constants.SHEPD.time_units)
			dfs[fault_time] = df

		# Combine results into a single DataFrame with an additional level to identify the fault by name.
		# Subsequent data extraction then deals with processing the relevant data
		self.df_combined_results = pd.concat(dfs.values(), axis=1, keys=dfs.keys())

		# Check for any negative R and X values and report busbars which have these values
		df_negative_impedance = self.df_combined_results[
			self.df_combined_results.loc[:, (slice(None), constants.BkdyFileOutput.x)] < 0
		].dropna(axis=0, how='all')
		if not df_negative_impedance.empty:
			negative_buses = df_negative_impedance.index
			self.unreliable_faulted_buses.extend(negative_buses)
			for bus in negative_buses:
				self.logger.warning(
					(
						'The busbar {} has a negative fault impedance value and therefore the fault current value '
						'returned by the PSSE BKDY method is unreliable and should not be used'
					).format(bus)
				)

		return self.df_combined_results

	def calculate_fault_currents(self, fault_times, g74_infeed, buses=list(), delete=True):
		"""
			Function calculates the fault currents at every busbar listed taking into consideration
			that the DC component and peak make has to be calculated based on t=0 and only the RMS
			symmetrical component should be recalculated to account for decrement.

			Two iterations of the fault current calculations are performed, one for every timestep with machines
			initialised for time == 0ms and then for every timestep with machine parameters recalculated.
		:param list fault_times:  List of the fault times that should be considered
		:param G74FaultInfeed() g74_infeed:  Reference to the g74 handle so that machine parameters can be updated
		:param list buses: (optional) List of busbars to be faulted if empty list then all busbars faulted
		:param bool delete: (optional=True) - Will delete the original bkdy output files
		:return None:
		"""
		# Fault current calculation to determine Ik'', peak make and DC decrement
		# Calculate the fault impedance values for the initial time of 0.0
		g74_infeed.calculate_machine_impedance(fault_time=0.0, update=True)

		# Initial fault must be carried out at 0.0 ms to get peak and Ik'' value
		if constants.G74.min_fault_time not in fault_times:
			fault_times.append(constants.G74.min_fault_time)
			self.logger.warning(
				(
					'{:.2f} fault time missing from inputs.  This must be included to determine the '
					'initial fault current.  This has been added to the fault'
					'times'
				).format(constants.G74.min_fault_time)
			)

		if constants.G74.peak_fault_time not in fault_times:
			fault_times.append(constants.G74.peak_fault_time)
			self.logger.warning(
				(
					'{:.2f} fault time missing from inputs.  This must be included to determine the '
					'peak current value in line with G74.  This time has been added to the fault'
					'times'
				).format(constants.G74.peak_fault_time)
			)

		# Sort list of times into ascending order
		fault_times.sort()

		# Produce name of results files for initial run
		# TODO: Change this to use the temporary folder rather than script folder (same folder as BKDY and log file outputs)
		current_script_path = os.path.dirname(os.path.realpath(__file__))
		initial_fault_files = [
			os.path.join(current_script_path, 'fault_ik_init{:.5f}{}'.format(x, constants.General.ext_csv))
			for x in fault_times
		]
		ac_decrement_files = [
			os.path.join(current_script_path, 'fault_ik_decr{:.5f}{}'.format(x, constants.General.ext_csv))
			for x in fault_times
		]

		# Define bus subsystem based on buses
		if buses:
			self.psse.define_bus_subsystem(buses=buses)
			self.sid = self.psse.sid
			self.logger.debug('Following busbars defined for fault analysis {}'.format(buses))
			self.all_buses = 0
		else:
			self.logger.info('No busbars defined and so all busbars will be faulted')
			self.all_buses = 1

		# Loop through fault current studies producing fault files initially for ik'' and DC component decay
		for fault, file_path in zip(fault_times, initial_fault_files):
			# Run fault study for this result
			# Fault is given name value for subsequent processing
			_t = time.time()
			self.logger.info(
				'Calculating fault current {:.2f} after fault application to determine DC decay'.format(fault)
			)
			self.main(name=fault, output_file=file_path, fault_time=fault)
			self.logger.info(
				'Fault currents {:.2f} seconds after application completed in {:.2f} seconds'.format(fault, time.time()-_t)
			)

		# Process results from initial fault into a DataFrame and delete if necessary
		df = self.combine_bkdy_output(delete=delete)

		# Loop through fault current studies producing fault files initially for ik(t)
		for fault, file_path in zip(fault_times, ac_decrement_files):
			# Recalculate machine parameters based on fault time
			g74_infeed.calculate_machine_impedance(fault_time=fault, update=True)
			# TODO: Make this capable as part of debugging for every fault time
			# Run fault study for this result
			_t = time.time()
			self.logger.info(
				(
					'Calculating fault current {:.2f} after fault application to determine reduced AC component'
				).format(fault)
			)
			self.main(name=fault, output_file=file_path, fault_time=fault)
			self.logger.info(
				(
					'Fault currents {:.2f} seconds after application completed in {:.2f} seconds'
				).format(fault, time.time() - _t)
			)

		# Process results from ik(t) fault into a DataFrame and delete results files if necessary
		df_decr = self.combine_bkdy_output(delete=delete)

		# Update ik(t) values in initial calculation with values from second DataFrame
		df.update(df_decr.xs(constants.BkdyFileOutput.ibsym, axis=1, level=1, drop_level=False))

		df = self.process_combined_results(df)
		df = self.add_busbar_data(df)
		return df

	def process_combined_results(self, df):
		"""
			Function will loop through and process the complete set of results to produce the data that is
			necessary for presenting
		:param pd.DataFrame() df:
		:return:
		"""
		#
		self.logger.debug('Combining results')
		dfs = dict()
		# TODO: How to extract section of DataFrame at this point
		for fault_time, df_section in df.groupby(level=0, axis=1):
			# Extract relevant sections
			if round(fault_time, 3) == constants.G74.min_fault_time:
				df_temp = df_section[fault_time][constants.SHEPD.cols_for_min_fault_time]
				# Calculate X/R value
				df_temp[constants.General.x_r] = (
					df_temp[constants.BkdyFileOutput.x].div(
						df_temp[constants.BkdyFileOutput.r]
					)
				)
			elif round(fault_time, 3) == constants.G74.peak_fault_time:
				df_temp = df_section[fault_time][constants.SHEPD.cols_for_peak_fault_time]
			else:
				df_temp = df_section[fault_time][constants.SHEPD.cols_for_other_fault_time]

			# Re-calculate asymmetrical fault current based on Iasym = sqrt(DC**2+((sqrt(2)SYM)**2)/2)
			df_temp[constants.BkdyFileOutput.ibasym] = (
				df_temp[constants.BkdyFileOutput.ibsym].mul(2**0.5).pow(2).div(2) +
				df_temp[constants.BkdyFileOutput.idc].pow(2)
			).pow(0.5)

			# Produce a name for this set of results which includes the fault time in the desired units
			name = '{} {}'.format(fault_time, constants.SHEPD.time_units)
			# Add to DataFrame dictionary
			dfs[name] = df_temp

		df = pd.concat(dfs.values(), axis=1, keys=dfs.keys(), names=constants.SHEPD.output_headers)
		df.sort_index(axis=1, level=0, inplace=True, ascending=True)
		return df

	def add_busbar_data(self, df):
		"""
			Function will add in busbar data to DataFrame
		:param pd.DataFrame df:
		:return:
		"""
		self.logger.debug('Adding busbar data to DataFrame')
		# Constants
		c = constants.General
		# Get busbar data from PSSE model
		bus_data = BusData()
		# Populate new DAtaFrame with relevant technical data based on indexes of busbars already faulted
		df_bus_data = pd.DataFrame(index=df.index)
		df_bus_data.loc[:, c.bus_name] = bus_data.df.loc[:, bus_data.c.bus_name]
		df_bus_data.loc[:, c.bus_voltage] = bus_data.df.loc[:, bus_data.c.nominal]
		df_bus_data.loc[:, c.pre_fault] = bus_data.df.loc[:, bus_data.c.voltage]
		# Convert to MultiIndex
		df_bus_data.columns = pd.MultiIndex.from_product(
			[[c.node_label], df_bus_data.columns],
			names=constants.SHEPD.output_headers
		)

		# Merge DataFrames together
		df = pd.concat([df_bus_data, df], axis=1)
		df.index.name = c.bus_number

		return df


# TODO: To be completed
class FormatResults:
	"""
		Convert the results into the specified format
	"""
	def __init__(self, psse, df, output_type='LTDS'):
		"""
		:param PsseControl psse:
		:param pd.DataFrame df:
		"""
		self.logger = logging.getLogger(constants.Logging.logger_name)

		self.psse = psse
		self.df = df
		self.output_type = output_type

		if output_type == 'LTDS':
			self.format_for_ltds()

	def format_for_ltds(self):
		"""
			Convert base df into format necessary for LTDS
		:return:
		"""
		self.logger.critical('Code not developed for LTDS format yet')
		return None
		# #bus_data = BusData()
		# #df_ltds = self.df
		# #pass
		# #df_ltds[bus_name] = None


class LoadData:
	"""
		Class that obtains all the data for the loads in the PSSE model
	"""

	def __init__(self, flag=1, sid=-1):
		"""
		:param int flag:  (optional=1) - Only returns details for loads at in-service busbars
		:param int sid:
		"""
		self.sid = sid
		self.flag = flag
		self.logger = logging.getLogger(constants.Logging.logger_name)

		self.c = constants.Loads

		self.df = pd.DataFrame()

		# Populate DataFrame
		self.update()

	def update(self):
		"""
			Update DataFrame with the data necessary for the idev file
		:return None:
		"""
		# Declare functions
		func_int = psspy.aloadint
		func_real = psspy.aloadreal
		func_char = psspy.aloadchar

		# Retrieve data from PSSE
		ierr_int, iarray = func_int(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.bus,))
		ierr_real, rarray = func_real(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.load,))
		ierr_char, carray = func_char(
			sid=self.sid,
			flag=self.flag,
			string=(self.c.identifier,))

		if ierr_int > 0 or ierr_char > 0 or ierr_real > 0:
			self.logger.critical(
				(
					'Unable to retrieve the load data from the SAV case and PSSE returned the '
					'following error codes {}, {} and {} from the functions <{}>, <{}> and <{}>'
				).format(ierr_int, ierr_real, ierr_char, func_int.__name__, func_real.__name__, func_char.__name__)
			)
			raise SyntaxError('Error importing data from PSSE SAV case')

		# Combine data into single list of lists
		data = iarray + rarray + carray
		# Column headers initially in same order as data but then reordered to something more useful for exporting
		initial_columns = [self.c.bus, self.c.load, self.c.identifier]

		# Transposed so columns in correct location and then columns reordered to something more suitable
		df = pd.DataFrame(data).transpose()
		df.columns = initial_columns

		self.df = df

		return None

	def summary(self):
		"""
			Produces a summary of the total load connected by busbar
		:return pd.DataFrame df_summary:
		"""
		# Extract the total load connected at each busbar and resulting index will be based on the busbar number
		df_summary = self.df.loc[:, (self.c.bus, self.c.load)].groupby(by=self.c.bus).sum(axis=1)
		# Adjust to only include data for those loads that are greater than zero
		df_summary = df_summary[df_summary[self.c.load] > constants.G74.min_load_mva]

		return df_summary


# TODO: Process output results to extract relevant values (input option to select values?)
class BkdyFile:
	def __init__(self, output_file, fault_time):
		"""
		:param str output_file:  Full path to output file that was produced by BKDY routine
		:param float fault_time:  Time of breaker separation for this study
		"""
		self.logger = logging.getLogger(constants.Logging.logger_name)
		# Define constants and initialise DataFrame
		self.output_file = output_file
		self.fault_time = fault_time

		# Will contain processed results
		self.df = pd.DataFrame()

	def process_bkdy_output(self, delete=False):
		"""
			Reads in the bkdy file and processes into a suitable DataFrame format
		:param bool delete:  (optional=False) - If set to True then will delete the file
		:return pd.DataFrame df:  DataFrame of all results in the file with column labels as listed in
								constants.BkdyFileOutput
		"""
		# Check if file has already been deleted and if so return previously imported and processed results
		if self.output_file is None:
			self.logger.error(
				(
					'Attempted to process a BKDY output file for fault time {:.2f} that has already file that has '
					'already been deleted.  Instead the previously imported and processed results will be returned.'
				).format(self.fault_time)
			)
			# Check if DataFrame is already empty which would be due to an issue in processing of the files
			if self.df.empty:
				self.logger.critical(
					'Something has gone wrong, attempted to process a file which has already '
					'been deleted but has not previously been processed.  Scripting error!'
				)
				raise SyntaxError('BKDY output file already deleted or empty')
			else:
				return self.df
		regex_bus = re.compile('[0-9]+')
		bus = int()
		start_reached = False
		c_bkdy_file = constants.BkdyFileOutput()
		with open(self.output_file, 'rb') as f:
			for line in f:
				# Find start of file
				if not start_reached and constants.BkdyFileOutput.start not in line:
					continue
				elif constants.BkdyFileOutput.start in line:
					start_reached = True
					continue

				# Find busbar number
				bus_line = regex_bus.search(line)
				if bus_line and not bus:
					bus = int(bus_line.group())
				elif constants.BkdyFileOutput.current in line:
					# Get relevant column numbers for this line
					col_nums, expected_length = c_bkdy_file.col_positions(line_type=c_bkdy_file.current)
					# Split the line into a list of floats
					currents = extract_values(line, expected_length)

					# Process results into DataFrame
					for name, col_num in col_nums.iteritems():
						self.df.loc[bus, name] = currents[col_num] / c_bkdy_file.num_to_kA

				elif constants.BkdyFileOutput.impedance in line:
					# TODO: Confirm base value of model to ensure values are presented on 100 MVA base
					# Get relevant column numbers for this line
					col_nums, expected_length = c_bkdy_file.col_positions(line_type=c_bkdy_file.impedance)
					# Split the line into a list of floats
					impedance = extract_values(line, expected_length=expected_length)

					# Process results into DataFrame
					for name, col_num in col_nums.iteritems():
						if col_num > 3:
							self.df.loc[bus, name] = impedance[col_num] / c_bkdy_file.num_to_kA
						else:
							self.df.loc[bus, name] = impedance[col_num]

					# Reset bus since finished processing this busbar
					bus = int()

		# Set name for DataFrame
		self.df.name = constants.BkdyFileOutput.start

		# Determine maximum values for DC and Peak current
		# TODO: Review if this is best method and not overly pessimistic with requirements of G74
		self.df[c_bkdy_file.ip] = self.df[[c_bkdy_file.ip_method1, c_bkdy_file.ip_method2]].max(axis=1)
		self.df[c_bkdy_file.idc] = self.df[[c_bkdy_file.idc_method1, c_bkdy_file.idc_method2]].max(axis=1)

		# Tidy up by removing file and updating status
		if delete:
			os.remove(self.output_file)
			self.output_file = None

		return self.df


class G74FaultInfeed:
	"""
		Class contains functions necessary for adding the equivalent fault in feeds to PSSE for asynchronous machines
		embedded as part of the load (LV and HV).  These machines are added inline with the G74 requirements.
	"""
	def __init__(self):
		"""
		"""
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.c = constants.G74
		# #self.df_hv_machines = hv_machines

		# Will contain details of all the loads in the model so that machines can be added
		self.df_load_buses = pd.DataFrame()
		# Will contain details of all the machines that need to be added
		self.df_machines = pd.DataFrame()
		# DataFrame will contain all the busbar data
		self.bus_data = pd.DataFrame()
		self.plant_data = pd.DataFrame()

		# Parameter set to True once machines have been added and checked
		self.machines_checked = False

	def identify_machine_parameters(self, hv_machines=pd.DataFrame()):
		"""
			Obtains details of all the loads in the system at each busbar along with the nominal voltage
		:param pd.DataFrame hv_machines:  DataFrame of any HV connected machines provided as an input, machines will be
									added at these busbars based on the HV parameters

		:return None:
		"""
		# Get load data and busbar data
		self.df_machines = LoadData().summary()
		self.bus_data = BusData()
		self.plant_data = PlantData()

		# Create DataFrame with details of machines that need to be added
		# Obtain nominal voltage from the busbar data
		self.df_machines[constants.Busbars.nominal] = self.bus_data.df[constants.Busbars.nominal]

		# Set flags for HV and LV machines accordingly (assume all lv to start with)
		self.df_machines[self.c.label_voltage] = self.c.lv
		self.df_machines[self.c.label_mva] = self.c.mva_11

		# If any HV machines have been added then parameters for these machines will also be added to the
		# output data and these will be based on the HV machine values
		# TODO: Reconsider approach here, values overwritten in calculate_machine_mva_values.  Though it may still be
		# TODO: useful to have an input available that can override the assumed parameters.
		if not hv_machines.empty:
			# TODO: Add in error checking for the case where HV machines are listed but there is no equivalent load
			hv_machines[self.c.label_voltage] = self.c.hv
			hv_machines[self.c.label_mva] = self.c.mva_hv
			self.df_machines[self.c.label_voltage].update(hv_machines[self.c.label_voltage])
			self.df_machines[self.c.label_mva].update(hv_machines[self.c.label_mva])

			# Check for any HV machines that are not in the DataFrame and report an error
			missing_hv_loads = hv_machines.loc[~hv_machines.index.isin(self.df_machines.index)]
			if not missing_hv_loads.empty:
				msg0 = (
					'The following HV connected loads have been listed as an input but no load is modelled in the PSSE '
					'case.  Therefore no equivalent infeed has been added for these:'
				)
				msg1 = '\n'.join([
					'\t - HV connected load at busbar {} has no equivalent load in PSSE case'.format(bus)
					for bus in missing_hv_loads.index
				])
				self.logger.warning('{}\n{}'.format(msg0, msg1))

		# Convert MVA values to be the values based on the load values
		# TODO: This no longer does anything since overwritten in calculate_machine_mva_values
		self.df_machines[self.c.label_mva] = self.df_machines[self.c.label_mva] * self.df_machines[constants.Loads.load]

		return None

	def calculate_machine_mva_values(self):
		"""
			Calculates the relevant values for the machines taking into consideration nominal voltage and transformer
			impedance data.  The X'', X' and X values are calculated based on the
		:return None:
		"""

		# Check if any loads modelled at > 33kV and display an error message
		df_above_33 = self.df_machines[self.df_machines[constants.Busbars.nominal] > 33.0]
		if not df_above_33.empty:
			msg0 = (
				'The following loads are modelled at greater than 33 kV and the ENA G74 guidance does not cover '
				'these.  They have been assumed to be modelled with the same parameters as applied for 33 kV '
				'equivalent fault infeed:'
			)
			msg1 = '\n'.join([
				'\t {:.2f} MVA load connected at busbar <{}> with nominal voltage {:.1f} kV'
				.format(mac[constants.Loads.load], bus, mac[constants.Busbars.nominal])
				for bus, mac in df_above_33.iterrows()]
			)
			self.logger.warning('{}\n{}'.format(msg0, msg1))

		# 33kV equivalent load parameters are as per ENA G74 technical data
		# Series created to allow easy adding to DataFrame
		# TODO: Efficiency improvement possible here since could just filter rows and assign in place
		df_33 = self.df_machines.loc[self.df_machines[constants.Busbars.nominal] > 11.0].assign(**self.c.parameters_33)
		df_11 = self.df_machines.loc[self.df_machines[constants.Busbars.nominal] <= 11.0].assign(**self.c.parameters_11)

		# Calculate MVA values for machine taking into consideration SHETL parameters
		df_33[self.c.label_mva] = self.c.mva_33 * df_33[constants.Loads.load]
		df_11[self.c.label_mva] = self.c.mva_11 * df_11[constants.Loads.load]

		# Combine back into a single data_frame
		self.df_machines = pd.concat([df_33, df_11], axis=0)

		self.logger.debug('Parameters calculated for machines connecting to represent embedded load at 11 and 33kV')

	def calculate_machine_impedance(self, fault_time, update=False):
		"""
			Calculates and updates the machine impedance values based on the fault time
		:param float fault_time: (optional=0.0) - X'', X' and X parameters based on the fault time input in seconds
		:param bool update:  If set to True then it will automatically update the machine impedance values once calculated
		:return None:
		"""
		# Calculate X'', X' and X values based on fault_time (based on equation 9.5.2 of G74 1992
		if fault_time > constants.PSSE.min_fault_time:
			x_value = 1.0 / ((1.0 / self.c.x11) * math.exp(-fault_time / self.c.t11))
		else:
			x_value = self.c.x11

		# Update DataFrame with these values
		c = constants.Machines
		# TODO: Confirm, this makes the assumption that the transformer impedance varies with the size of
		# TODO: the load connected which doesn't seem fully correct.
		# Update values based on new x_value taking into consideration the transformer reactance
		self.df_machines.loc[:, (
									c.xsubtr,
									c.xtrans,
									c.xsynch
								)] = x_value

		self.df_machines.loc[:, (
									c.xsubtr,
									c.xtrans,
									c.xsynch
								)] = self.df_machines.loc[:, (
									c.xsubtr,
									c.xtrans,
									c.xsynch
								)].subtract(self.df_machines.loc[:, c.tx_x], axis=0)

		self.logger.debug(
			(
				"G74 machine values updated for a fault time of {:.2f} seconds based on an x'' of {:.3f} p.u., "
				"time constant of {:.2f} seconds.  Resulting in x at time of fault of {:.3f} p.u."
			).format(fault_time, self.c.x11, self.c.t11, x_value))

		# If set to True then will automatically go and update the machine impedance values once calculated
		if update:
			self.add_machines()

	def add_machines(self):
		"""
			Adds / updates the parameters for every machine in the PSSE base case to ensure the G74 contribution
			is included.  Will also change the state of busbars to generator buses where appropriate.
		:return None:
		"""
		func_machine = psspy.machine_data_2
		func_machine_seq = psspy.seq_machine_data_3
		func_bus = psspy.bus_data_3
		func_plant = psspy.plant_data

		# Loop through every machine and add / update parameters in the PSSE case
		for bus, machine in self.df_machines.iterrows():
			# Check busbar state is the correct type (type codes 2, 3 or 4 do not impact)
			# Must be done before adding machine otherwise get a missing Plant Data error
			if self.bus_data.df.loc[bus, constants.Busbars.state] == 1:
				# If busbar is type code 1 (non-generator bus) then change status to 2
				ierr_bus = func_bus(
					i=bus,
					intgar1=constants.Busbars.generator_bus_type_code
				)
			else:
				ierr_bus = 0

			# Check if plant already exists and if not add Plant
			if bus not in self.plant_data.df.loc[:, constants.Plant.bus].tolist():
				ierr_plant = func_plant(
					i=bus
				)
			else:
				ierr_plant = 0

			# Add machine / update MVA values
			# TODO: label_mva is not recognised and so is returning 0 (need to check where this should be populated from)
			ierr_mac = func_machine(
				i=bus,
				id=self.c.machine_id,
				intgar1=1,			# Ensures machine is in service
				realar1=0.0,		# Ensures machine P output is 0.0 (PG)
				realar2=0.0,		# Ensures machine Q output is 0.0 (QG)
				realar3=0.0,		# Ensures machine Q output is 0.0 (QT)
				realar4=0.0,		# Ensures machine Q output is 0.0 (QB)
				realar5=0.0,		# Ensures machine P output is 0.0 (PT)
				realar6=0.0,		# Ensures machine P output is 0.0 (PB)
				realar7=machine[self.c.label_mva],
				realar8=machine[constants.Machines.rsource],
				realar9=machine[constants.Machines.xsource]
			)

			# Update machine sequence values
			ierr_seq = func_machine_seq(
				i=bus,
				id=self.c.machine_id,
				realar1=machine[constants.Machines.rpos],
				realar2=machine[constants.Machines.xsubtr],
				realar3=machine[constants.Machines.rneg],
				realar4=machine[constants.Machines.xneg],
				realar5=machine[constants.Machines.rzero],
				realar6=machine[constants.Machines.xzero],
				realar7=machine[constants.Machines.xtrans],
				realar8=machine[constants.Machines.xsynch]
			)

			# Error checking / debug writing
			if sum([ierr_mac, ierr_seq, ierr_bus, ierr_plant]) > 0:
				self.logger.error(
					(
						'An error occurred when trying to add an equivalent machine to represent the fault current '
						'contribution from embedded load to the busbar {}.  The functions <{}>, <{}>, <{}> and <{}> '
						'returned the following error codes: {}, {}, {} and {}'
					).format(
						bus,
						func_bus.__name__, func_plant, func_machine.__name__, func_machine_seq.__name__,
						ierr_bus, ierr_plant, ierr_mac, ierr_seq
					)
				)
			else:
				self.logger.debug(
					'Machine parameters successfully updated for equivalent machine connected to busbar: {} with ID {}'
					.format(bus, constants.G74.machine_id)
				)


class IecFaults:
	"""
		Class for carrying out IEC fault current calculations and returning the required data
	TODO:  Need to implement an initial stage to determine the DC component and the peak component
	TODO: Validate IEC method for 3Ph and LG conforms with G74
	"""
	def __init__(self, psse, buses=list()):
		"""
			List of busbars to consider for IEC fault current calculations
		:param PsseControl psse:  Controller for psse
		:param list buses: (optional=list)
		"""
		self.logger = logging.getLogger(constants.Logging.logger_name)

		# Check if bus subsystem needs defining
		if len(buses) == 0:
			psse.sid = -1
		else:
			psse.define_bus_subsystem(buses=buses)

		self.buses = buses
		self.sid = psse.sid

		# Values used for processing results
		self.bus_data = BusData()
		self.result_unit = str()
		self.result_coordinate = str()

		# Set fault units and coordinates to the correct formats
		psse.set_outputs()

	def extract_value(self, value_to_convert, bus):
		"""
			Function processes an individual result to extract the relevant format based on the output format
		:param complex value_to_convert:  Value that needs returning in the relevant format
		:param int bus:  Relevant busbar number for this bus, required if extracting the unit data
		:return float value: The converted value that is returned
		"""
		if self.result_coordinate == 'rectangular':
			# If rectangular then magnitude is given by absolute of complex number
			value = abs(value_to_convert)
		elif self.result_coordinate == 'polar':
			# If polar then first value is magnitude and second value is angle
			value = value_to_convert.real
		else:
			self.logger.critical(
				'Unexpected value <{}> returned for IEC fault current coordinates'.format(self.result_coordinate)
			)
			raise SyntaxError('Unexpected value returned for IEC fault current coordinates')

		# Convert to required kA or A value
		if self.result_unit == 'pu':
			bus_nominal_voltage = self.bus_data.df.loc[bus, self.bus_data.c.nominal]
			# Convert value to kA
			value = value*(constants.PSSE.base_mva / (bus_nominal_voltage*3**0.5))
		elif self.result_unit == 'physical':
			value = value / constants.BkdyFileOutput.num_to_kA
		else:
			self.logger.critical(
				'Unexpected value <{}> returned for IEC fault current results unit'.format(self.result_unit)
			)
			raise SyntaxError('Unexpected value returned for IEC fault current unit')

		return value

	def fault_3ph_all_buses(self, fault_time):
		"""
			Calculate three phase fault using IEC methodology
		:param float fault_time:  Breaker opening time
		:return pd.DataFrame df:
		"""
		# PSSE functions
		func_iecs = pssarrays.iecs_currents

		# Get latest busbar data
		self.bus_data = BusData()
		# If looking at all busbars then produce list of buses based on all busbars
		if self.sid == -1:
			buses_to_fault = self.bus_data.df[self.bus_data.c.bus].tolist()
		else:
			buses_to_fault = self.buses

		# Constant definition for all output data
		c = constants.BkdyFileOutput
		df = pd.DataFrame()

		# Loop through each busbar and perform fault current calculation
		for bus in buses_to_fault:
			# Get the pre-fault voltage for this busbar
			pre_fault_v = self.bus_data.df.loc[bus, self.bus_data.c.voltage]
			# TODO: Need to confirm parameters for IEC fault current calculation
			iec_results = func_iecs(
				sid=self.sid,
				flt3ph=1,
				fltloc=0,
				# Line charging set to 1, 0.0 in positive and negative sequences
				lnchrg=1,
				# Zero sequence transformer impedance correction is ignored
				zcorec=0,
				# Load treated as 0.0 in positive, negative sequences
				loadop=1,
				optnftrc=2,
				brktime=fault_time,
				vfactorc=pre_fault_v
			)
			if iec_results.ierr > 0:
				self.logger.critical(
					(
						'Error running a fault current calculation on busbar {} with pre fault voltage of {:.2f} and '
						'circuit breaker opening time of {:.2f} seconds.  The function <{}> returned the error code {} '
					).format(bus, pre_fault_v, fault_time, func_iecs.__name__, iec_results.ierr)
				)
				raise ValueError('Error running IEC fault current')
			else:
				self.result_coordinate = iec_results.scfmt
				self.result_unit = iec_results.scunit

				bus_idx = iec_results.fltbus.index(bus)
				df.loc[bus, c.ik11] = self.extract_value(iec_results.flt3ph[bus_idx].ia1, bus)
				df.loc[bus, c.ip] = self.extract_value(iec_results.flt3ph[bus_idx].ipc, bus)
				df.loc[bus, c.idc] = self.extract_value(iec_results.flt3ph[bus_idx].idc, bus)
				df.loc[bus, c.ibsym] = self.extract_value(iec_results.flt3ph[bus_idx].ibsym, bus)
				df.loc[bus, c.ibasym] = self.extract_value(iec_results.flt3ph[bus_idx].ibuns, bus)

		# Return the DataFrame of the results for 3 phase faults
		return df
