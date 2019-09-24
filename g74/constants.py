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

import os

# Set to True to run in debug mode and therefore collect all output to window
DEBUG_MODE = True

fault_time = 0.06

class SHEPD:
	"""
		Constants specific to the WPD study
	"""
	# voltage_limits are declared as a dictionary in the format
	# {(lower_voltage,upper_voltage):(pu_limit_lower,pu_limit_upper)} where
	# <lower_voltage> and <upper_voltage> represent the extremes over which
	# the voltage <pu_limit_lower> and <pu_limit_upper> applies
	# These are based on the post-contingency steady state limits provided in
	# EirGrid, "Transmission System Security and Planning Standards"
	# 380 used since that is base voltage in PSSE
	steady_state_limits = {
		(109, 111.0): (99.0/110.0, 120.0/110.0),
		(219.0, 221.0): (200.0/220.0, 240.0/220.0),
		(250.0, 276.0): (250.0/275.0, 303.0/275.0),
		(379.0, 401.0): (360.0/380.0, 410.0/380.0)
	}

	reactor_step_change_limit = 0.03
	cont_step_change_limit = 0.1

	# This is a threshold value, circuits with ratings less than this are reported and ignored
	rating_threshold = 0

	# Default time constant values to assume
	t1d0 = 0.04
	t11d0 = 0.12
	t1q0 = t1d0
	t11q0 = t11d0

	def __init__(self):
		""" Purely added to avoid error message"""
		pass


class PSSE:
	"""
		Class to hold all of the constants associated with PSSE initialisation
	"""
	# default version = 33
	version = 33

	# Setting on whether PSSE should output results based on whether operating in DEBUG_MODE or not
	output = {True: 1, False: 6}

	# Maximum number of iterations for a Newton Raphson load flow (default = 20)
	max_iterations = 100
	# Tolerance for mismatch in MW/Mvar (default = 0.1)
	mw_mvar_tolerance = 1.0

	sid = 1

	# Load Flow Constants
	tie_line_flows = 0,  # Don't enable tie line flows
	phase_shifting = 0,  # Phase shifting adjustment disabled
	dc_tap_adjustment = 0,  # DC tap adjustment disabled
	var_limits = 0,  # Apply VAR limits immediately
	non_divergent = 0

	def __init__(self):
		self.psse_py_path = str()
		self.psse_os_path = str()

	def get_psse_path(self, psse_version, reset=False):
		"""
			Function returns the PSSE path specific to this version of psse
		:param int psse_version:
		:param bool reset: (optional=False) - If set to True then class is reset with a new psse_version
		:return str self.psse_path:
		"""
		if self.psse_py_path and self.psse_os_path and not reset:
			return self.psse_py_path, self.psse_os_path

		if 'PROGRAMFILES(X86)' in os.environ:
			program_files_directory = r'C:\Program Files (x86)\PTI'
		else:
			program_files_directory = r'C:\Program Files\PTI'

		psse_paths = {
			32: 'PSSE32\PSSBIN',
			33: 'PSSE33\PSSBIN',
			34: 'PSSE34\PSSPY27'
		}
		os_paths = {
			32: 'PSSE32\PSSBIN',
			33: 'PSSE33\PSSBIN',
			34: 'PSSE34\PSSBIN'
		}
		self.psse_py_path = os.path.join(program_files_directory, psse_paths[psse_version])
		self.psse_os_path = os.path.join(program_files_directory, os_paths[psse_version])
		return self.psse_py_path, self.psse_os_path

	def find_psspy(self, start_directory=r'C:'):
		"""
			Function to search entire directory and find PSSE installation

		"""
		# Clear variables
		self.psse_py_path = str()
		self.psse_os_path = str()

		psspy_to_find = "psspy.pyc"
		psse_to_find = "psse.bat"
		for root, dirs, files in os.walk(start_directory):  # Walks through all subdirectories searching for file
			if psspy_to_find in files:
				self.psse_py_path = root
			elif psse_to_find in files:
				self.psse_os_path = root

			if self.psse_py_path and self.psse_os_path:
				break

		return self.psse_py_path, self.psse_os_path


class Machines:
	bus = 'NUMBER'
	identifier = 'ID'
	x_synch = 'XSYNCH'
	x_trans = 'XTRANS'
	x_subtr = 'XSUBTR'

	t1d0 = "T'd0"
	t11d0 = "T''d0"
	t1q0 = "T'q0"
	t11q0 = "T''q0"

	xd = 'Xd'
	xq = 'Xq'
	x1d = "X'd"
	x1q = "X'q"
	x11 = "X''"

	bkdy_col_order = [bus, identifier, t1d0, t1q0, t11d0, t11q0, xd, xq, x1d, x1q, x11]

	def __init__(self):
		pass


class Busbars:
	bus = 'NUMBER'
	state = 'TYPE'
	nominal = 'BASE'
	voltage = 'PU'
	bus_name = 'EXNAME'

	# Labels used for columns to determine voltage limits
	nominal_lower = 'NOM_LOWER'
	nominal_upper = 'NOM_UPPER'
	lower_limit = 'LOWER_LIMIT'
	upper_limit = 'UPPER_LIMIT'

	# Number stored in DataFrame if an error has occurred
	error_id = -1

	def __init__(self):
		pass


class Logging:
	"""
		Log file names to use
	"""
	logger_name = 'JK7938'
	debug = 'DEBUG'
	progress = 'INFO'
	error = 'ERROR'
	extension = '.log'

	def __init__(self):
		"""
			Just included to avoid Pycharm error message
		"""
		pass


class Excel:
	""" Constants associated with inputs from excel """
	circuit = 'Circuits'
	tx2 = '2 Winding'
	tx3 = '3 Winding'
	busbars = 'Busbars'
	fixed_shunts = 'Fixed Shunts'
	switched_shunts = 'Switched Shunts'
	machine_data = 'Machines'

	def __init__(self):
		pass
