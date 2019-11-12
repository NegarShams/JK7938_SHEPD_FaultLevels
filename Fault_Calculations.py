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

# This import needs to be run first to ensure that the script path is defined correctly in which some other
# modules will be located (numpy and pandas in particular) that are required as part of this script package
import os
script_path = os.path.realpath(__file__)
script_folder = os.path.dirname(script_path)

import g74
import g74.constants as constants
import time
try:
	import pandas as pd
except ImportError, err:
	# TODO: Write function that will go and install the missing files using pip and try again
	raise ImportError(
		(
			'Unable to import pandas most likely because it has not been installed correctly within the'
			'folder that constains this script!\n'
			'It should be installed manually using <pip> and the appropriate <.whl> using the command:\n'
			'\t<pip install --ignore-installed --no-deps --target="{}" "{}\<name of file>.whl"\n'
			'The full traceback was as follows:\n{}'
		).format(script_folder, script_folder, err)
	)

# If set to True then will delete all raw results data
DELETE_RESULTS = False


def bkdy_study(
		sav_case, local_temp_folder, excel_file, export_sav_case,
		fault_times, buses, local_logger, reload_sav=True
):
	"""
		Run BKDY G74 fault study calculation
	:param str sav_case:
	:param str local_temp_folder:
	:param str excel_file:
	:param str export_sav_case:
	:param list fault_times:  Times that fault study is run for
	:param list buses:  List of busbars to fault
	:param g74.Logger local_logger:
	:param bool reload_sav:  Whether original SAV case should be reloaded at the end
	:return:
	"""
	t1 = time.time()
	temp_bkd_file = os.path.join(local_temp_folder, 'bkdy_machines{}'.format(constants.PSSE.ext_bkd))

	# Load SAV case
	psse = g74.psse.PsseControl()
	psse.load_data_case(pth_sav=sav_case)

	local_logger.app = psse
	print('Running from PSSE status is: {}'.format(logger.app.run_in_psse))
	local_logger.info('Running from PSSE status is: {}'.format(logger.app.run_in_psse))

	t2 = time.time() - t1
	t1 = time.time()

	# Carry out BKDY calculation
	bkdy = g74.psse.BkdyFaultStudy(psse_control=psse)
	bkdy.create_breaker_duty_file(target_path=temp_bkd_file)

	t3 = time.time() - t1
	t1 = time.time()

	# Update model to include contribution from embedded machines
	g74_data = g74.psse.G74FaultInfeed()
	g74_data.identify_machine_parameters()
	g74_data.calculate_machine_mva_values()

	t4 = time.time() - t1
	t1 = time.time()

	# Save temporary SAV case
	psse.save_data_case(pth_sav=export_sav_case)

	df = bkdy.calculate_fault_currents(
		fault_times=fault_times, g74_infeed=g74_data,
		# #buses=buses_to_fault, delete=False
		buses=buses,
		delete=True
	)

	# Save temporary SAV case
	psse.save_data_case(pth_sav=export_sav_case)

	t5 = time.time() - t1
	t1 = time.time()

	# Export results to excel
	with pd.ExcelWriter(path=excel_file) as writer:
		df.to_excel(writer, sheet_name='Fault I')
		df.T.to_excel(writer, sheet_name='Fault I Transposed')
	local_logger.info('Excel workbook written to: {}'.format(excel_file))
	t6 = time.time() - t1

	# Export tabulated data to PSSE
	# TODO: Decide if this would be useful and how to format

	# Will reload original SAV case if required
	if reload_sav:
		psse.load_data_case(pth_sav=pth_sav_case)
		local_logger.debug('Original sav case: {} reloaded'.format(pth_sav_case))

	# Restore output to defaults
	psse.change_output(destination=1)

	return t2, t3, t4, t5, t6


if __name__ == '__main__':
	# Create logger
	uid = 'BKDY_{}'.format(time.strftime('%Y%m%d_%H%M%S'))
	# Check temp folder exists and if not create
	temp_folder = os.path.join(script_folder, 'temp')
	if not os.path.exists(temp_folder):
		os.mkdir(temp_folder)
	logger = g74.Logger(pth_logs=temp_folder, uid=uid, debug=constants.DEBUG_MODE)

	t0 = time.time()
	# Run main study
	logger.info('Study started')
	# Load user interface
	gui = g74.gui.MainGUI()

	# Whether SAV case should be reloaded
	reload_sav_case = gui.bo_reload_sav.get()

	# TODO: Get SAV case from GUI and also initially save it so that it can be reloaded on completion

	pth_sav_case = gui.sav_case
	faults = gui.fault_times
	target_file = gui.target_file
	sav_name, _ = os.path.splitext(os.path.basename(pth_sav_case))
	pth_sav_case_export = os.path.join(temp_folder, '{}.sav'.format(sav_name))
	buses_to_fault = gui.selected_busbars

	times = bkdy_study(
		sav_case=pth_sav_case, local_temp_folder=temp_folder, excel_file=target_file,
		export_sav_case=pth_sav_case_export, fault_times=faults, buses=buses_to_fault,
		reload_sav=reload_sav_case, local_logger=logger
	)

	logger.info('Took {:.2f} seconds to initialise PSSE'.format(times[0]))
	logger.info('Took {:.2f} seconds to create BKDY files for machines'.format(times[1]))
	logger.info('Took {:.2f} seconds to add G74 machines'.format(times[2]))
	logger.info('Took {:.2f} seconds to carry out fault study and save cases'.format(times[3]))
	logger.info('Took {:.2f} seconds to export to excel'.format(times[4]))

	logger.info('Complete with total study time of {:.2f} seconds'.format(time.time()-t0))
