"""
#######################################################################################################################
###											PSSE G7/4 Fault Studies													###
###		Test to ensure Python and PSSE install is working correctly													###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

# Assumes that system comes with capability of producing logging messages
import logging
import os
import time
import sys


class InitialisePsspy:
	"""
		Class to deal with the initialising of PSSE by checking the correct directory is being referenced and has been
		added to the system path and then attempts to initialise it
	"""
	def __init__(self, psse_version=33):
		"""
			Initialise the paths and checks that import psspy works
		:param int psse_version: (optional=34)
		"""

		self.psse = False

		# Get PSSE path
		self.psse_py_path, self.psse_os_path = self.get_psse_path(psse_version=psse_version)
		# Add to system path if not already there
		if self.psse_py_path not in sys.path:
			sys.path.append(self.psse_py_path)

		if self.psse_os_path not in os.environ['PATH']:
			os.environ['PATH'] += ';{}'.format(self.psse_os_path)

		if self.psse_py_path not in os.environ['PATH']:
			os.environ['PATH'] += ';{}'.format(self.psse_py_path)

		global psspy
		global pssarrays
		try:
			# Import psspy used for manipulating PSSE
			import psspy
			psspy = reload(psspy)
			self.psspy = psspy
			# Import pssarrays used for data extraction from PSSE
			import pssarrays
			pssarrays = reload(pssarrays)
			self.pssarrays = pssarrays
			# #import pssarrays
			# #self.pssarrays = pssarrays
		except ImportError:
			self.psspy = None
			# #self.pssarrays = None

	def get_psse_path(self, psse_version):
		"""
			Function returns the PSSE path specific to this version of psse
		:param int psse_version:
		:return str self.psse_path:
		"""
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
				self.change_output(destination=1)

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


def setup_logger():
	"""
		Function sets up an error logger to record messages
	:return Logging.logger logger:
	"""
	# Get path to script
	script_path = os.path.dirname(os.path.realpath(__file__))
	# Setup logging
	# logging.getLogger().disabled = True
	_logger = logging.getLogger()
	# Ensures that even debug messages are captured
	_logger.setLevel(logging.DEBUG)
	# Produce formatter for log entries
	log_formatter = logging.Formatter(
		fmt='%(asctime)s - %(levelname)s - %(message)s',
		datefmt='%Y-%m-%d %H:%M:%S')
	# Create file handler
	handler = logging.FileHandler(filename=os.path.join(script_path, 'log_python_tester.txt'))
	handler.setFormatter(log_formatter)
	_logger.addHandler(handler)

	return _logger


if __name__ == '__main__':
	# Initialise and create logger
	t0 = time.time()
	t1 = time.time()
	logger = setup_logger()
	logger.info('Logger created in {:.2f} now testing module imports'.format(time.time()-t1))
	t1 = time.time()

	# Test imports
	count = 0
	try:
		import numpy as np
		logger.debug('numpy imported')
	except ImportError:
		logger.error('Unable to import <numpy>')
		count += 1

	try:
		import pandas as pd
		logger.debug('pandas imported')
	except ImportError:
		logger.error('Unable to import <pandas>')
		count += 1

	try:
		import re
		logger.debug('re imported')
	except ImportError:
		logger.error('Unable to import <re>')
		count += 1

	try:
		import math
		logger.debug('math imported')
	except ImportError:
		logger.error('Unable to import math')
		count += 1

	try:
		import string
		logger.debug('string imported')
	except ImportError:
		logger.error('Unable to import string')
		count += 1

	# Report update
	if count > 0:
		logger.warning('Imports tested and there were {} import failures'.format(count))

	logger.info('Imports completed in {:.2f} seconds'.format(time.time()-t1))
	t1 = time.time()

	# Initialise PSSE
	success = InitialisePsspy().initialise_psse()

	if success:
		logger.debug('PSSE successfully initialised')
	else:
		logger.error('Unable to initialise PSSE')
		count += 1

	if count > 0:
		logger.error('Errors during test, review log messages')
	else:
		logger.info('All tests ran successfully in {:.2f} seconds'.format(time.time()-t0))
