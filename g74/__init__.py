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

# Add local_packages in parent to path
# Generic Imports
import os
import sys
import logging
import logging.handlers
import time
import inspect
import subprocess

# TODO: May be better to actually embed this within G74
local_packages = os.path.join(os.path.dirname(__file__), '..', 'local_packages')
# Won't be searched unless it exists when added to system path
if not os.path.exists(local_packages):
	os.makedirs(local_packages)
# Insert local_packages to start of path for fault studies
sys.path.insert(0, local_packages)

# Package imports
try:
	import g74.constants as constants
	import g74.psse as psse
	import g74.file_handling as file_handling
	import g74.gui as gui
except ImportError:
	t0 = time.time()
	# TODO: Add in a check to confirm that the files actually exist
	print(
		'Unable to import some packages because they may not have been installed, script will now install'
		'missing packages but this may take some time, please be patient!!'
	)
	batch_path = os.path.join(os.path.dirname(__file__), '..', 'setup_SHEPD.bat')
	print('The following batch file will be run to install the packages: {}'.format(batch_path))
	subprocess.call([batch_path])
	print(
		(
			'Unless the batch file showed an error packages have now been installed and took {:.2f} seconds'
		).format(time.time()-t0)
	)
	import g74.constants as constants
	import g74.psse as psse
	import g74.file_handling as file_handling
	import g74.gui as gui
	print('All modules now imported correctly')

# Meta Data
__author__ = 'David Mills'
__version__ = '0.1'
__email__ = 'david.mills@PSCconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'Development'


def decorate_emit(fn):
	"""
		Function will decorate the log message to insert a colour control for the console output
	:param fn:
	:return:
	"""

	# add methods we need to the class
	def new(*args):
		level_number = args[0].levelno
		if level_number >= logging.CRITICAL:
			# Set colour to red
			color = '\x1b[31;1m'
		elif level_number >= logging.ERROR:
			# Set colour to Red
			color = '\x1b[31;1m'
		elif level_number >= logging.WARNING:
			# Set colour to dark yellow
			color = '\x1b[33;1m'
		elif level_number >= logging.INFO:
			# Set colour to yellow
			color = '\x1b[32;1m'
		elif level_number >= logging.DEBUG:
			# Set colour to purple
			color = '\x1b[35;1m'
		else:
			color = '\x1b[0m'

		# Change the colour of the log messages
		args[0].msg = "{0}{1}\x1b[0m ".format(color, args[0].msg)
		args[0].levelname = "{0}{1}\x1b[0m ".format(color, args[0].levelname)

		return fn(*args)

	return new


class Logger:
	"""
		Customer logger for dealing with log output during script runs
	"""
	def __init__(self, pth_logs, uid, app=None, debug=False):
		"""
			Initialise logger
		:param str pth_logs:  Path to where all log files will be stored
		:param str uid:  Unique identifier for log files
		:param bool debug:  True / False on whether running in debug mode or not
		:param g74.psse.PsseControl() app: (optional) - If not None then will use this to provide updates to powerfactory
		"""
		# Constants
		self.log_constants = constants.Logging

		# Attributes used during setup_logging
		self.handler_progress_log = None
		self.handler_debug_log = None
		self.handler_error_log = None
		self.handler_stream_log = None

		# Counter for each error message that occurs
		self.warning_count = 0
		self.error_count = 0
		self.critical_count = 0

		# Populate default paths
		self.pth_logs = pth_logs
		self.pth_debug_log = os.path.join(pth_logs, 'DEBUG_{}.log'.format(uid))
		self.pth_progress_log = os.path.join(pth_logs, 'INFO_{}.log'.format(uid))
		self.pth_error_log = os.path.join(pth_logs, 'ERROR_{}.log'.format(uid))
		self.app = app
		self.debug_mode = debug

		self.file_handlers = []

		# Set up logger and establish handle for logger
		self.check_file_paths()
		self.logger = self.setup_logging()
		self.initial_log_messages()

	def check_file_paths(self):
		"""
			Function to check that the file paths are accessible
		:return None:
		"""
		script_pth = os.path.realpath(__file__)
		parent_pth = os.path.abspath(os.path.join(script_pth, os.pardir))
		uid = time.strftime('%Y%m%d_%H%M%S')

		# Check each file to see if it can be created or if it even exists, if not then use script directory
		if self.pth_debug_log is None:
			file_name = '{}_{}{}'.format(self.log_constants.debug, uid, self.log_constants.extension)
			self.pth_debug_log = os.path.join(parent_pth, file_name)

		if self.pth_progress_log is None:
			file_name = '{}_{}{}'.format(self.log_constants.progress, uid, self.log_constants.extension)
			self.pth_progress_log = os.path.join(parent_pth, file_name)

		if self.pth_error_log is None:
			file_name = '{}_{}{}'.format(self.log_constants.error, uid, self.log_constants.extension)
			self.pth_progress_log = os.path.join(parent_pth, file_name)

		return None

	def setup_logging(self):
		"""
			Function to setup the logging functionality
		:return object logger:  Handle to the logger for writing messages
		"""
		# logging.getLogger().setLevel(logging.CRITICAL)
		# logging.getLogger().disabled = True
		logger = logging.getLogger(self.log_constants.logger_name)
		logger.handlers = []

		# Ensures that even debug messages are captured even if they are not written to log file
		logger.setLevel(logging.DEBUG)

		# Produce formatter for log entries
		log_formatter = logging.Formatter(
			fmt='%(asctime)s - %(levelname)s - %(message)s',
			datefmt='%Y-%m-%d %H:%M:%S')

		self.handler_progress_log = self.get_file_handlers(
			pth=self.pth_progress_log, min_level=logging.INFO, _buffer=True, flush_level=logging.ERROR,
			formatter=log_formatter)

		self.handler_debug_log = self.get_file_handlers(
			pth=self.pth_debug_log, min_level=logging.DEBUG, _buffer=True, flush_level=logging.CRITICAL,
			buffer_cap=100000, formatter=log_formatter)

		self.handler_error_log = self.get_file_handlers(
			pth=self.pth_error_log, min_level=logging.ERROR, formatter=log_formatter)

		self.handler_stream_log = logging.StreamHandler()

		# If running in DEBUG mode then will export all the debug logs to the window as well
		self.handler_stream_log.setFormatter(log_formatter)
		if self.debug_mode:
			self.handler_stream_log.setLevel(logging.DEBUG)
		else:
			self.handler_stream_log.setLevel(logging.INFO)

		# Decorate to colour code different warning labels
		self.handler_stream_log.emit = decorate_emit(self.handler_stream_log.emit)

		# Add handlers to logger
		logger.addHandler(self.handler_progress_log)
		logger.addHandler(self.handler_debug_log)
		logger.addHandler(self.handler_error_log)
		logger.addHandler(self.handler_stream_log)

		return logger

	def initial_log_messages(self):
		"""
			Display initial messages for logger including paths where log files will be stored
		:return:
		"""
		# Initial announcement of directories for log messages to be saved in
		self.info(
			'Path for debug log is {} and will be created if any WARNING messages occur'.format(self.pth_debug_log))
		self.info(
			'Path for process log is {} and will contain all INFO and higher messages'.format(self.pth_progress_log))
		self.info(
			'Path for error log is {} and will be created if any ERROR messages occur'.format(self.pth_error_log))
		self.debug(
			(
				'Stream output is going to stdout which will only be displayed if DEBUG MODE is True and currently it '
				'is {}').format(self.debug_mode)
		)

		# Ensure initial log messages are created and saved to log file
		self.handler_progress_log.flush()
		return None

	def close_logging(self):
		"""Function closes logging but first removes the debug_handler so that the output is not flushed on
			completion.
		"""
		# Close the debug handler so that no debug outputs will be written to the log files again
		# This is a safe close of the logger and any other close, i.e. an exception will result in writing the
		# debug file.
		# Flush existing progress and error logs
		self.handler_progress_log.flush()
		self.handler_error_log.flush()

		# Specifically remove the debug_handler
		self.logger.removeHandler(self.handler_debug_log)

		# Close and delete file handlers so no more logs will be written to file
		for handler in reversed(self.file_handlers):
			handler.close()
			del handler

	def get_file_handlers(self, pth, min_level, formatter, _buffer=False, flush_level=logging.INFO, buffer_cap=10):
		"""
			Function to a handler to write to the target file with our without a buffer if required
			Files are overwritten if they already exist
		:param str pth:  Path to the file handler to be used
		:param int min_level: Is the minimum level that the file handler should include
		:param bool _buffer: (optional=False)
		:param int flush_level: (optional=logging.INFO) - The level at which the log messages should be flushed
		:param int buffer_cap:  (optional=10) - Level at which the buffer empties
		:param logging.Formatter formatter:  (optional=logging.Formatter()) - Formatter to use for the log file entries
		:return: logging.handler handler:  Handle for new logging handler that has been created
		"""
		# Handler for process_log, overwrites existing files and buffers unless error message received
		# delay=True prevents the file being created until a write event occurs

		handler = logging.FileHandler(filename=pth, mode='a', delay=True)
		self.file_handlers.append(handler)

		# Add formatter to log handler
		handler.setFormatter(formatter)

		# If a buffer is required then create a new memory handler to buffer before printing to file
		if _buffer:
			handler = logging.handlers.MemoryHandler(
				capacity=buffer_cap, flushLevel=flush_level, target=handler)

		# Set the minimum level that this logger will process things for
		handler.setLevel(min_level)

		return handler

	def progress_output(self):
		"""
			Function toggles the GUI progress output if running in PSSE
		:return None:
		"""
		if self.app is not None:
			if self.app.run_in_psse:
				self.app.toggle_progress_output(destination=1)
		return None

	def no_progress_output(self):
		"""
			Function toggles the GUI progress output if running in PSSE
		:return None:
		"""
		if self.app is not None:
			if self.app.run_in_psse:
				self.app.toggle_progress_output(destination=6)
		return None

	def debug(self, msg):
		""" Handler for debug messages """
		# Debug messages only written to logger
		self.progress_output()
		self.logger.debug(msg)
		self.no_progress_output()

	def info(self, msg):
		""" Handler for info messages """
		# # Only print output to powerfactory if it has been passed to logger
		# #if self.app and self.pf_executed:
		# #	self.app.PrintPlain(msg)
		self.progress_output()
		self.logger.info(msg)
		self.no_progress_output()

	def warning(self, msg):
		""" Handler for warning messages """
		self.warning_count += 1
		# #if self.app and self.pf_executed:
		# #	self.app.PrintWarn(msg)
		self.progress_output()
		self.logger.warning(msg)
		self.no_progress_output()

	def error(self, msg):
		""" Handler for warning messages """
		self.error_count += 1
		# #if self.app and self.pf_executed:
		# #	self.app.PrintError(msg)
		self.progress_output()
		self.logger.error(msg)
		self.no_progress_output()

	def critical(self, msg):
		""" Critical error has occurred """
		# Get calling function to include in log message
		# https://stackoverflow.com/questions/900392/getting-the-caller-function-name-inside-another-function-in-python
		caller = inspect.stack()[1][3]
		self.critical_count += 1

		self.progress_output()
		self.logger.critical('function <{}> reported {}'.format(caller, msg))
		self.no_progress_output()

	def flush(self):
		""" Flush all loggers to file before continuing """
		self.handler_progress_log.flush()
		self.handler_error_log.flush()

	def logging_final_report_and_closure(self):
		"""
			Function reports number of error messages raised and closes down logging
		:return None:
		"""
		if sum([self.warning_count, self.error_count, self.critical_count]) > 1:
			self.logger.info(
				(
					'Log file closing, there were the following number of important messages: \n'
					'\t - {} Warning Messages that may be of concern\n'
					'\t - {} Error Messages that may have stopped the results being produced\n'
					'\t - {} Critical Messages').format(self.warning_count, self.error_count, self.critical_count)
			)
		else:
			self.logger.info('Log file closing, there were 0 important messages')
		self.logger.debug('Logging stopped')
		logging.shutdown()

	def __del__(self):
		"""
			To correctly handle deleting and therefore shutting down of logging module
		:return None:
		"""
		self.logging_final_report_and_closure()

	def __exit__(self):
		"""
			To correctly handle deleting and therefore shutting down of logging module
		:return None:
		"""
		self.logging_final_report_and_closure()
