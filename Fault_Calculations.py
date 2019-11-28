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

# Modules to be imported first
import os
import g74
import g74.constants as constants
import time
import pandas as pd

# Meta Data
__author__ = 'David Mills'
__version__ = '0.0.1'
__email__ = 'david.mills@PSCconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'In Development - Beta'


def fault_study(
		psse_handler,
		local_uid, sav_case, local_temp_folder, excel_file, fault_times, buses, local_logger, reload_sav=True
):
	"""
		Run G74 fault study calculation using PSSE BKDY or IEC methods and obtain the
		fault current at all the busbars that have been listed.
	:param g74.PsseControl psse_handler:  Handle for the psse interface engine
	:param str local_uid:  Unique identifier for this study used to append to files
	:param str sav_case:  Full path to the SAV case that should be used for the fault study
	:param str local_temp_folder:  Local temporary folder into which to save temporary data
	:param str excel_file:  Full path to where results should be exported
	:param list fault_times:  Times that fault study is run for
	:param list buses:  List of busbars to fault
	:param g74.Logger local_logger:  Path to logger
	:param bool reload_sav:  (optional=True) Whether original SAV case should be reloaded at the end
	:return None:
	"""
	# Produce temporary files
	t = time.time()
	temp_bkd_file = os.path.join(local_temp_folder, 'bkdy_machines{}'.format(constants.PSSE.ext_bkd))

	# Get path for export SAV case
	sav_name, _ = os.path.splitext(os.path.basename(sav_case))
	temp_sav_case = os.path.join(local_temp_folder, '{}_{}.sav'.format(sav_name, local_uid))

	# Initialise PSSE and load SAV case
	psse_handler.load_data_case(pth_sav=sav_case)

	# Get handle to logger and determine whether running for PSSE or from Python
	local_logger.app = psse_handler
	print('Running from PSSE status is: {}'.format(logger.app.run_in_psse))
	local_logger.info('Running from PSSE status is: {}'.format(logger.app.run_in_psse))
	local_logger.info('Took {:.2f} seconds to initialise PSSe and load SAV case'.format(time.time()-t))
	t = time.time()

	# Create the files for the existing machines that will be used for the BKDY fault study
	bkdy = g74.psse.BkdyFaultStudy(psse_control=psse_handler)
	bkdy.create_breaker_duty_file(target_path=temp_bkd_file)
	local_logger.info('Took {:.2f} seconds to create BKDY files for machines'.format(time.time()-t))
	t = time.time()

	# Update model to include contribution from embedded machines
	g74_data = g74.psse.G74FaultInfeed()
	g74_data.identify_machine_parameters()
	g74_data.calculate_machine_mva_values()
	local_logger.info(
		(
			'Took {:.2f} seconds to add G74 machines that represent contribution from embedded load'
		).format(time.time()-t)
	)
	t = time.time()

	# Save a temporary SAV case if necessary
	if temp_sav_case:
		psse_handler.save_data_case(pth_sav=temp_sav_case)

	# TODO:  At this point want to add in also IEC fault study for 3Ph and LG
	# Carry out fault current study for each time step
	df = bkdy.calculate_fault_currents(
		fault_times=fault_times, g74_infeed=g74_data,
		# #buses=buses_to_fault, delete=False
		buses=buses,
		delete=True
	)

	# Save temporary SAV case (if necessary)
	if temp_sav_case:
		psse_handler.save_data_case(pth_sav=temp_sav_case)
	local_logger.info('Took {:.2f} seconds to carry out all fault current studies.'.format(time.time() - t))
	t = time.time()

	# Export results to excel
	with pd.ExcelWriter(path=excel_file) as writer:
		df.to_excel(writer, sheet_name='Fault I')
		df.T.to_excel(writer, sheet_name='Fault I Transposed')
	local_logger.info('Results written to Excel workbook: {}'.format(excel_file))
	local_logger.info('Took {:.2f} seconds to save results'.format(time.time()-t))
	t = time.time()

	# Export tabulated data to PSSE
	# TODO: Export tabulated data directly to PSSE window

	# Will reload original SAV case if required
	if reload_sav:
		psse_handler.load_data_case(pth_sav=pth_sav_case)
		local_logger.debug('Original sav case: {} reloaded'.format(pth_sav_case))

	# Restore output to defaults
	psse_handler.change_output(destination=1)

	# Produce error message at end of output to report potential busbar fault error issues
	if bkdy.unreliable_faulted_buses:
		msg0 = (
			'The following busbars had an issue carrying out the fault current study which has been reported above and '
			'as such the value for these busbars is unreliable:'
		)
		msg1 = '\n'.join(['\t - {}'.format(bus) for bus in set(bkdy.unreliable_faulted_buses)])
		logger.warning('{}\n{}'.format(msg0, msg1))

	local_logger.info(
		'Took {:.2f} seconds to reload SAV case and export warning messages (if any)'.format(time.time()-t)
	)

	return None


def get_busbars(psse_handler):
	"""
		Determines if PSSE is running and if so will return a list of busbars
	:param g74.PsseControl psse_handler:  Handle to controller for psse
	:return:
	"""
	psse_handler.running_from_psse()
	if psse_handler.run_in_psse:
		# Initialise PSSE so have access to all required variables
		_ = g74.psse.InitialisePsspy().initialise_psse(running_from_psse=psse_handler.run_in_psse)
		busbars = g74.psse.PsseSlider().get_selected_busbars()
		sav_case = psse_handler.get_current_sav_case()
	else:
		busbars = list()
		sav_case = None

	return busbars, sav_case


if __name__ == '__main__':
	"""
		This is the main block of code that will be run if this script is run directly
	"""
	# Time stamp for performance checking
	t0 = time.time()

	# Produce unique identifier for logger
	uid = 'BKDY_{}'.format(time.strftime('%Y%m%d_%H%M%S'))

	# Check temp folder exists to store log files in and if not create appropriate folders
	script_path = os.path.realpath(__file__)
	script_folder = os.path.dirname(script_path)
	temp_folder = os.path.join(script_folder, 'temp')
	if not os.path.exists(temp_folder):
		os.mkdir(temp_folder)
	logger = g74.Logger(pth_logs=temp_folder, uid=uid, debug=constants.DEBUG_MODE)

	# Run main study
	logger.info('Study started')

	# Check if PSSE is running and if so retrieve list of selected busbars, else return empty list
	psse = g74.psse.PsseControl()
	selected_busbars, current_sav_case = get_busbars(psse)

	if current_sav_case:
		# TODO: Better way to do this, maybe GUI popup asking user whether to save SAV case before making updates
		logger.warning(
			'If the current SAV case has not been saved prior to running study, changes will be lost when reloaded'
		)

	# Load GUI and ask user to select required inputs
	gui = g74.gui.MainGUI(sav_case=current_sav_case, busbars=selected_busbars)

	# Determine whether user aborted study rather than selecting SAV case
	if gui.abort:
		logger.warning('User interface closed by user and study aborted after {:.2f} seconds'.format(time.time()-t0))
	else:
		# Get path to SAV case being faulted
		pth_sav_case = gui.sav_case
		# Whether SAV case should be reloaded
		reload_sav_case = gui.bo_reload_sav.get()

		# Get parameters from GUI
		faults = gui.fault_times
		target_file = gui.target_file
		buses_to_fault = gui.selected_busbars
		open_excel = gui.bo_open_excel.get()

		fault_study(
			psse_handler=psse,
			local_uid=uid, sav_case=pth_sav_case, local_temp_folder=temp_folder, excel_file=target_file,
			fault_times=faults, buses=buses_to_fault, reload_sav=reload_sav_case, local_logger=logger,
		)

		# Open the exported excel if setting is as such
		# TODO: Alternatively, adjust to just display in an instance of excel rather than having to save the results
		# TODO: Detect if already open and if so, warn user and save with a different name
		if open_excel:
			os.startfile(target_file)

		logger.info('Complete with total study time of {:.2f} seconds'.format(time.time()-t0))
