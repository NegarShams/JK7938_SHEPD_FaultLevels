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
import re

# Set to True to run in debug mode and therefore collect all output to window
DEBUG_MODE = False

# TODO: Change fault time to be either a GUI input or a constant
fault_time = 0.06
# TODO: Define as a constant input
convert_to_kA = True


class General:
	"""
		General constants
	"""
	ext_csv = '.csv'

	def __init__(self):
		"""
			Just to avoid error message
		"""
		pass


class GUI:
	"""
		Constants for the user interface
	"""
	gui_name = 'PSC G7/4 Fault Current Tool'

	def __init__(self):
		"""
			Purely to avoid error message
		"""
		pass


class BkdyFileOutput:
	"""
		Constants for processing BKDY file output
	"""
	base_mva = 100.0

	# TODO: Add this a constants value
	if convert_to_kA:
		num_to_kA = 1000.0
		current_unit = 'kA'
	else:
		num_to_kA = 1.0
		current_unit = 'A'

	start = 'FAULTED BUS'
	current = 'FAULT CURRENT'
	impedance = 'THEVENIN IMPEDANCE'

	ik11 = "Ik'' ({})".format(current_unit)
	ip = 'Ip ({})'.format(current_unit)
	ibsym = 'Ibsym ({})'.format(current_unit)
	ibasym = 'Ibasym ({})'.format(current_unit)
	idc = 'DC ({})'.format(current_unit)
	idc0 = 'DC_t0 ({})'.format(current_unit)
	v_prefault = 'V Pre-fault (p.u.)'

	# Impedance values
	x = 'X (p.u. on {:.0f} MVA)'.format(base_mva)
	r = 'R (p.u. on {:.0f} MVA)'.format(base_mva)

	# Error flag if Vpk returns infinity
	infinity_error = '*******'

	reg_search = re.compile('(\d\.\d+)|(\d+\.\d)|(-\d+\.\d{2})')

	# NaN value that is returned if error calculating fault current values
	nan_value = 'NaN'
	# This is replaced with the following and an error message given to user
	# TODO: Ensure error message is given to user
	nan_replacement = '0.0'

	def __init__(self):
		"""
			Purely to avoid error message
		"""
		pass

	def col_positions(self, line_type):
		"""
			Returns a dictionary with the associated column positions depending on the line type
		:param str line_type:  based on the values defined above returns the relevant column numbers
		:return dict, int (cols, expected_length):  Dictionary of column positions, expected length of list of floats
		"""
		cols = dict()
		if line_type == self.current:
			cols[self.ik11] = 0
			cols[self.ibsym] = 2
			cols[self.idc] = 4
			cols[self.ibasym] = 5
			cols[self.ip] = 6

			# Expected length of this list of floats
			expected_length = 7
		elif line_type == self.impedance:
			cols[self.r] = 0
			cols[self.x] = 1
			cols[self.v_prefault] = 2
			# Not possible to export this data since in some cases get a result returned which says infinity
			# #cols[self.idc0] = 4
			# #cols[self.ibasym0] = 5
			# #cols[self.ip0] = 6

			# Expected length of this list of floats
			expected_length = 7
		else:
			raise ValueError(
				(
					'The line_type <{}> provided does not match the available options of:\n'
					'\t - {}\n'
					'\r - {}\n'
					'Check the code!'
				).format(line_type, self.current, self.impedance)
			)

		return cols, expected_length


class PSSE:
	"""
		Class to hold all of the constants associated with PSSE initialisation
	"""
	# default version = 33
	version = 33

	# Setting on whether PSSE should output results based on whether operating in DEBUG_MODE or not
	output = {True: 1, False: 6}
	file_output = 2

	# Maximum number of iterations for a Newton Raphson load flow (default = 20)
	max_iterations = 100
	# Tolerance for mismatch in MW/Mvar (default = 0.1)
	mw_mvar_tolerance = 1.0

	sid = 1

	# Load Flow Constants
	tie_line_flows = 0  # Don't enable tie line flows
	phase_shifting = 0  # Phase shifting adjustment disabled
	dc_tap_adjustment = 0  # DC tap adjustment disabled
	var_limits = 0  # Apply VAR limits immediately
	non_divergent = 0

	ext_bkd = '.bkd'

	# Default parameters for PSSE outputs
	# 1 = polar coordinates
	def_short_circuit_units = 1
	# 1 = physical units
	def_short_circuit_coordinates = 1

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


class Loads:
	bus = 'NUMBER'
	load = 'MVAACT'
	identifier = 'ID'

	def __init__(self):
		"""
			Purely to avoid error codes
		"""
		pass


class Machines:
	bus = 'NUMBER'
	identifier = 'ID'
	rpos = 'RPOS'
	rneg = 'RNEG'
	rzero = 'RZERO'
	xsynch = 'XSYNCH'
	xtrans = 'XTRANS'
	xsubtr = 'XSUBTR'
	xneg = 'XNEG'
	xzero = 'XZERO'
	zsource = 'ZSORCE'
	rsource = 'R Source'
	xsource = 'X Source'

	t1d0 = "T'd0"
	t11d0 = "T''d0"
	t1q0 = "T'q0"
	t11q0 = "T''q0"

	xd = 'Xd'
	xq = 'Xq'
	x1d = "X'd"
	x1q = "X'q"
	x11 = "X''"

	# Minimum expected realistic RPOS value
	min_r_pos = 0.0
	# Assumed X/R value when they are missing
	assumed_x_r = 40.0

	bkdy_col_order = [bus, identifier, t1d0, t11d0, t1q0, t11q0, xd, xq, x1d, x1q, x11]

	# Defines the option for psspy.cong with regards to treatment of conventional machines and induction machines
	# 0 = Uses Zsorce for conventional machines
	# 1 = Uses X'' for conventional machines
	# 2 = Uses X' for conventional machines
	# 3 = Uses X for conventional machines
	bkdy_machine_type = 0

	def __init__(self):
		pass


class Plant:
	bus = 'NUMBER'
	status = 'STATUS'

	def __init__(self):
		"""
			Purely to avoid error messages
		"""
		pass


class Busbars:
	bus = 'NUMBER'
	state = 'TYPE'
	nominal = 'BASE'
	voltage = 'PU'
	bus_name = 'EXNAME'

	# Busbar type code lookup
	generator_bus_type_code = 2

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


class G74:
	# Assumed X/R ratio of equivalent motor connected at 33kV
	x_r_33 = 2.76
	# MVA contribution of equivalent motor per MVA of connected load (some ratio of these may be needed
	# based on whether load is assumed to be LV or HV connected.
	# TODO: Determine ratios of LV and HV connected load
	label_mva = 'Machine Base'
	mva_lv = 1.0
	mva_hv = 2.6

	# 11 and 33kV parameters as per SHETL documentation
	# TODO: Validate SHETL parameters and document in report
	mva_33 = 1.16
	mva_11 = 2.3

	# Minimum MVA value for load to be considered
	min_load_mva = 0.15

	# Labels for DataFrame
	label_voltage = 'Load Voltage'
	hv = 'hv'
	lv = 'lv'

	machine_id = 'LD'

	# Time constants
	t11 = 0.04
	# ## TODO: Confirm this time constant is sufficient, should have decayed to 0 my 120 ms
	# #t2 = 0.12

	# Calculation of R and X'' for equivalent machine connected at 33kV and assumes
	# Z=1.0 which is then multiplied by the MVA rating of the machine
	rpos = (1.0/(1.0+x_r_33**2))**0.5
	# #rneg = rpos
	x11 = (1.0-rpos**2)**0.5
	# #xneg = x11
	# x1 and X calculated as per G74 equations
	# #x1 = 1.0/((1.0/x11)*math.exp(-fault_time/t1))
	# #x1 = 1.0/((1.0/x11)*math.exp(-t1/t1))
	# ## Very high value for X1 and R0, X0 to ensure decays to 0
	# #x = 10000.0
	rzero = 10000.0
	xzero = 10000.0

	# Convert parameters to dictionary for easy updating in PSSe
	parameters_33 = {
			Machines.rpos: rpos,
			Machines.xsubtr: x11,
			Machines.xtrans: x11,
			Machines.xsynch: x11,
			Machines.rneg: rpos,
			Machines.xneg: x11,
			Machines.rzero: rzero,
			Machines.xzero: xzero,
			Machines.xsource: x11,
			Machines.rsource: rpos
		}

	# TODO: Calculate parameters for 33/11kV transformers and sensitivity study to determine the impact of these values
	# Transformer data in per unit on 100MVA base values
	# No longer accounting for 33/11kV transformers on a case by case basis but applying
	# SHETL parameters detailed above
	# #tx_r = 0.07142
	# #tx_x = 1.0

	def __init__(self):
		"""
			Purely to avoid error message
		"""
		pass


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
		(109, 111.0): (99.0 / 110.0, 120.0 / 110.0),
		(219.0, 221.0): (200.0 / 220.0, 240.0 / 220.0),
		(250.0, 276.0): (250.0 / 275.0, 303.0 / 275.0),
		(379.0, 401.0): (360.0 / 380.0, 410.0 / 380.0)
	}

	reactor_step_change_limit = 0.03
	cont_step_change_limit = 0.1

	# This is a threshold value, circuits with ratings less than this are reported and ignored
	rating_threshold = 0

	# Default time constant values to assume
	t1d0 = 0.12
	t11d0 = 0.04
	t1q0 = t1d0
	t11q0 = t11d0

	# Names and results associated with each type of result
	cb_make = 'Make'
	cb_break = 'Break'
	cb_steady = 'Steady'
	results_per_fault = dict()
	results_per_fault[cb_make] = [
		BkdyFileOutput.ik11,
		BkdyFileOutput.ip,
		BkdyFileOutput.x,
		BkdyFileOutput.r
	]
	results_per_fault[cb_break] = [
		BkdyFileOutput.ibsym,
		BkdyFileOutput.ibasym
	]

	# TODO: This should be defined as an input but the default position is these values
	fault_times = [0.01, fault_time]

	# List controls the order of the output columns for the LTDS export
	output_column_order = [
		BkdyFileOutput.ik11,
		BkdyFileOutput.ip,
		BkdyFileOutput.ibsym,
		BkdyFileOutput.r,
		BkdyFileOutput.x
	]

	def __init__(self):
		""" Purely added to avoid error message"""
		pass


class LTDS:
	bus_name = 'Name'
	nominal = 'Voltage (kV)'

	def __init__(self):
		""" Purely added to avoid error message"""
		pass
