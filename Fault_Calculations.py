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

import g74
import g74.constants as constants
import os
import time

parent_dir = (
	r'C:\Users\david\Power Systems Consultants Inc\Jobs - JK7938 - Automated Systems & '
	r'Data Process Efficiencies\5 Working Docs\PSSE Fault Calculation'
)
TEMP_FOLDER = os.path.join(parent_dir, 'temp')
SAV_NAME = 'SHEPD 2018 LTDS Winter Peak(all_gen_on)'
PSSE_SAV_CASE = os.path.join(parent_dir, '{}.sav'.format(SAV_NAME))
PSSE_SAV_CASE_EXPORT = os.path.join(TEMP_FOLDER, '{}.sav'.format(SAV_NAME))
EXCEL_FILE = os.path.join(parent_dir, '{}.xlsx'.format(SAV_NAME))

def bkdy_study(sav_case, temp_folder, excel_file, export_sav_case):
	"""
		Run BKDY G74 fault study calculation
	:param str sav_case:
	:param str temp_folder:
	:param str excel_file:
	:param str export_sav_case:
	:return:
	"""
	t1 = time.time()
	temp_bkd_file = os.path.join(temp_folder, 'bkdy_machines{}'.format(constants.PSSE.ext_bkd))
	fault_times = constants.SHEPD.fault_times

	output_files = list()
	for t in fault_times:
		output_files.append(os.path.join(temp_folder, 'bkdy_output{:.2f}{}'.format(t, constants.General.ext_csv)))
	names = (constants.SHEPD.cb_make, constants.SHEPD.cb_break)

	# Load SAV case
	psse = g74.psse.PsseControl()
	psse.load_data_case(pth_sav=sav_case)
	t2 = time.time() - t1

	t1 = time.time()

	# Update model to include contribution from embedded machines
	g74_data = g74.psse.G74FaultInfeed()
	g74_data.identify_machine_parameters()
	g74_data.calculate_machine_mva_values()

	# Carry out BKDY calculation
	bkdy = g74.psse.BkdyFaultStudy(psse_control=psse)
	bkdy.create_breaker_duty_file(target_path=temp_bkd_file)

	g74_data.add_machines()

	# #psse.save_data_case(pth_sav=export_sav_case)
	# #raise SyntaxError('STOP')

	t3 = time.time()-t1
	t1 = time.time()

	# Run fault study for each fault time
	for f, flt_time, name in zip(output_files, fault_times, names):
		bkdy.main(output_file=f, fault_time=flt_time, name=name)

	psse.save_data_case(pth_sav=export_sav_case)
	t6 = time.time() - t1
	t1 = time.time()

	# Process output files into DataFrame
	df = bkdy.combine_bkdy_output()
	t4 = time.time() - t1

	t1 = time.time()

	# Reformat df to match with required output
	df = df[constants.SHEPD.output_column_order]
	# Export results to excel
	df.to_excel(excel_file)
	t5 = time.time() - t1

	return t2, t3, t4, t5, t6


if __name__ == '__main__':
	# Create logger
	uid = 'BKDY_{}'.format(time.strftime('%Y%m%d_%H%M%S'))
	logger = g74.Logger(pth_logs=TEMP_FOLDER, uid=uid, debug=constants.DEBUG_MODE)

	t0 = time.time()
	# Run main study
	logger.info('Study started')
	times = bkdy_study(
		sav_case=PSSE_SAV_CASE, temp_folder=TEMP_FOLDER, excel_file=EXCEL_FILE,
		export_sav_case=PSSE_SAV_CASE_EXPORT
	)

	logger.info('Took {:.2f} seconds to initialise PSSE'.format(times[0]))
	logger.info('Took {:.2f} seconds to add G74 machines'.format(times[1]))
	logger.info('Took {:.2f} seconds to carry out BKDY calculation'.format(times[2]))
	logger.info('Took {:.2f} seconds to export to excel'.format(times[3]))
	logger.info('Took {:.2f} seconds to sav final case'.format(times[4]))

	logger.info('Complete with total study time of {:.2f} seconds'.format(time.time()-t0))
