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
import Tkinter as tk
import tkFileDialog

import sys
import os
import logging
import g74
import g74.constants as constants


class MainGUI:
	"""
		Main class to produce the GUI
		Allows the user to select the busbars and methodology to be applied in the fault current calculations
	"""
	def __init__(self, title=constants.GUI.gui_name, sav_case=str()):
		"""
			Initialise GUI
		:param str title: (optional) - Title to be used for main window
		:param str sav_case: (optional) - Full path to the existing SAV case
		TODO: Add option to reload initial SAV case on completion (default to ticked)
		"""
		# Get logger handle
		self.logger = logging.getLogger(constants.Logging.logger_name)

		# Initialise constants and tk window
		self.master = tk.Tk()
		self.master.title(title)

		self.fault_times = list()
		# General constants which need to be initialised
		self._row = 0
		self._col = 0

		# Will be populated with a list of busbars to be faulted
		self.selected_busbars = list()
		# Target file that results will be exported to
		self.target_file = str()
		self.results_pth = os.path.dirname(os.path.realpath(__file__))

		# General constants
		self.cmd_select_sav_case = tk.Button()
		self.cmd_import_busbars = tk.Button()
		self.cmd_edit_busbars = tk.Button()
		self.var_fault_times_list = tk.StringVar()
		self.entry_fault_times = tk.Entry()
		self.cmd_calculate_faults = tk.Button()
		self.bo_reload_sav = tk.BooleanVar()
		self.bo_fault_3_ph_bkdy = tk.BooleanVar()
		self.bo_fault_3_ph_iec = tk.BooleanVar()
		self.bo_fault_1_ph_iec = tk.BooleanVar()
		self.bo_open_excel = tk.BooleanVar()

		# SAV case for faults to be run on
		self.sav_case = sav_case

		# Add button for selecting SAV case to run
		self.add_cmd_sav_case(row=self.row(1), col=self.col())
		self.add_reload_sav(row=self.row(1), col=self.col())

		_ = tk.Label(self.master, text='Select Fault Studies to Include').grid(
			row=self.row(1), column=self.col(), sticky=tk.W+tk.E
		)
		# Add tick boxes for fault types to include
		self.add_fault_types(col=self.col())

		# Add button for importing / viewing busbars
		self.add_cmd_import_busbars(row=self.row(1), col=self.col())
		self.add_cmd_edit_busbars(row=self.row(), col=self.col()+1)

		# Add a text entry box for fault times to be added as a comma separated list
		self.add_entry_fault_times(row=self.row(1), col=self.col())

		# Add button for calculating and saving fault currents
		self.add_cmd_calculate_faults(row=self.row(2), col=self.col())

		# Add tick box for whether it needs to be opened again on completion
		self.add_open_excel(row=self.row(1), col=self.col())

		self.logger.debug('GUI window created')
		# Produce GUI window
		self.master.mainloop()

	def row(self, i=0):
		"""
			Returns the current row number + i
		:param int i: (optional=0) - Will return the current row number + this value
		:return int _row:
		"""
		self._row += i
		return self._row

	def col(self, i=0):
		"""
			Returns the current col number + i
		:param int i: (optional=0) - Will return the current col number + this value
		:return int _row:
		"""
		self._col += i
		return self._col

	def add_cmd_sav_case(self, row, col):
		"""
			Function just adds the command button to the GUI which is used for selecting the SAV case
		:param int row: Row number to use
		:param int col: Column number to use
		:return None:
		"""
		# Determine label for button based on the SAV case being loaded
		if self.sav_case:
			lbl_sav_button = 'SAV case = {}'.format(os.path.basename(self.sav_case))
		else:
			lbl_sav_button = 'Select SAV Case for Fault Study'

		# Create button and assign to Grid
		self.cmd_select_sav_case = tk.Button(
			self.master,
			text=lbl_sav_button,
			command=self.select_sav_case)
		self.cmd_select_sav_case.grid(row=row, column=col, columnspan=2, sticky=tk.W+tk.E)
		CreateToolTip(widget=self.cmd_select_sav_case, text=(
			'Select the SAV case for which fault studies should be run.'
		))
		return None

	def add_cmd_import_busbars(self, row, col):
		"""
			Function just adds the command button to the GUI which is used for selecting a spreadsheet with
			a list of busbars to fault
		:param int row: Row number to use
		:param int col: Column number to use
		:return None:
		"""
		# Add button for selecting busbars to fault
		lbl_button = 'Import list of busbars'
		self.cmd_import_busbars = tk.Button(
			self.master,
			text=lbl_button,
			command=self.import_busbars_list)
		self.cmd_import_busbars.grid(row=row, column=col)
		CreateToolTip(widget=self.cmd_import_busbars, text=(
			'Import a list of busbars from a spreadsheet for further editing.'
		))
		return None

	def add_cmd_edit_busbars(self, row, col):
		"""
			Function pops up a window to allow editing of the busbar list
		:param int row: Row number to use
		:param int col: Column number to use
		:return None:
		"""
		# Add button for selecting busbars to fault
		lbl_button = 'Edit Busbars List'
		self.cmd_edit_busbars = tk.Button(
			self.master,
			text=lbl_button,
			command=self.edit_busbars_list)
		self.cmd_edit_busbars.grid(row=row, column=col)
		self.cmd_edit_busbars.config(state='disabled')
		CreateToolTip(widget=self.cmd_edit_busbars, text=(
			'A new window will popup which allows the busbars list to be edited / reviewed\n'
			'TODO: Code has not been written for this yet'
		))
		return None

	def add_entry_fault_times(self, row, col):
		"""
			Function to add the text entry row for inserting fault times
		:param int row:  Row number to use
		:param int col:  Column number to use
		:return None:
		"""
		# Label for what is included in entry
		lbl = tk.Label(master=self.master, text='Desired Fault Times\n(in seconds separated by commas)')
		lbl.grid(row=row, column=col, rowspan=2, sticky=tk.W+tk.N+tk.S)
		# Set initial value for variable
		self.var_fault_times_list.set(constants.GUI.default_fault_times)
		# Add entry box
		self.entry_fault_times = tk.Entry(master=self.master, textvariable=self.var_fault_times_list)
		self.entry_fault_times.grid(row=row, column=col + 1, sticky=tk.W+tk.E, rowspan=2)
		CreateToolTip(widget=self.entry_fault_times, text=(
			'Enter the durations after the fault the current should be calculated for.\n'
			'Multiple values can be input in a list.'
		))
		return None

	def add_cmd_calculate_faults(self, row, col):
		"""
			Function to add the command button for completing the GUI entry and calculating fault currents
		:param int row:  Row number to use
		:param int col:   Column number to use
		:return None:
		"""
		lbl = 'Run Fault Study'
		# Add command button
		self.cmd_calculate_faults = tk.Button(
			self.master, text=lbl, command=self.process
		)
		self.cmd_calculate_faults.grid(row=row, column=col, columnspan=2, sticky=tk.W+tk.E)
		CreateToolTip(widget=self.cmd_calculate_faults, text='Calculate fault currents')
		return None

	def add_reload_sav(self, row, col):
		"""
			Function to add a tick box on whether the user wants to reload the SAV case
		:param int row:  Row number to use
		:param int col:  Column number to use
		:return None:
		"""
		lbl = 'Reload initial SAV case on completion'
		self.bo_reload_sav.set(constants.GUI.reload_sav_case)
		# Add tick box
		check_button = tk.Checkbutton(
			self.master, text=lbl, variable=self.bo_reload_sav
		)
		check_button.grid(row=row, column=col, columnspan=2, sticky=tk.W)
		CreateToolTip(widget=check_button, text=(
			'If selected the SAV case will be reloaded at the end of this study, if not then the model will as the '
			'study finished which may be useful for debugging purposes.'
		))
		return None

	def add_fault_types(self, col):
		"""
			Function to add a tick box to select the available fault types
		:param int col:  Column number to use
		:return None:
		"""
		labels = (
			'3 Phase fault (BKDY method)',
			'3 Phase fault (IEC method)',
			'LG Phase fault (IEC method)'
		)
		boolean_vars = (self.bo_fault_3_ph_bkdy, self.bo_fault_3_ph_iec, self.bo_fault_1_ph_iec)
		enabled = (
			True,
			False,
			False
		)
		for i, lbl in enumerate(labels):
			# Defaults assuming that all faults will be calculated
			boolean_vars[i].set(1)
			# Add check button for this fault
			check_button = tk.Checkbutton(
				self.master, text=lbl, variable=boolean_vars[i]
			)
			check_button.grid(row=self.row(1), column=col, sticky=tk.W)
			# TODO: Temporary to disable non-necessary faults
			# Disable faults that are not important
			if not enabled[i]:
				check_button.config(state='disabled')
		return i

	def add_open_excel(self, row, col):
		"""
			Function to add a tick box on whether the user wants to open the Excel file of results at the end
			:param int row:  Row number to use
			:param int col:  Column number to use
			:return None:
		"""
		lbl = 'Open exported Excel file'
		self.bo_open_excel.set(constants.GUI.open_excel)
		# Add tick box
		check_button = tk.Checkbutton(
			self.master, text=lbl, variable=self.bo_reload_sav
		)
		check_button.grid(row=row, column=col, columnspan=2, sticky=tk.W)
		CreateToolTip(widget=check_button, text=(
			'If selected the exported excel file will be loaded and visible on completion of the study.'
		))
		return None

	def import_busbars_list(self):
		"""
			Function to import a list of busbars based on the selected file
		:return: None
		"""
		# Ask user to select file(s) or folders based on <.bo_files>
		file_path = tkFileDialog.askopenfilename(
			initialdir=self.results_pth,
			filetypes=constants.General.file_types,
			title='Select spreadsheet containing list of busbars'
		)

		# Import busbar list from file assuming it is first column and append to existing list
		busbars = g74.file_handling.import_busbars_list(path=file_path)
		self.selected_busbars.extend(busbars)

		# Update results path to include this name
		self.results_pth = os.path.dirname(file_path)

		# Enable button to allow popup of busbars to be edited to be created
		# TODO: Still to be created
		self.cmd_edit_busbars.config(state='disabled')
		return None

	def edit_busbars_list(self):
		"""
			Function to popup a list of busbars that are to be faulted so that new / different busbars can be
			added or removed from the list
		:return None:
		"""
		# TODO: Write a pop-up window that will allow busbars list to be edited / manually populated
		pass

	def select_sav_case(self):
		"""
			Function to allow the user to select the SAV case to run
		:return: None
		"""
		# Ask user to select file(s) or folders based on <.bo_files>
		file_path = tkFileDialog.askopenfilename(
			initialdir=self.results_pth,
			filetypes=constants.General.sav_types,
			title='Select SAV case for fault studies'
		)

		# Import busbar list from file assuming it is first column and append to existing list
		self.sav_case = file_path
		lbl_sav_button = 'SAV case = {}'.format(os.path.basename(self.sav_case))

		# Update command button
		self.cmd_select_sav_case.config(text=lbl_sav_button)
		return None

	def process(self):
		"""
			Function sorts the files list to remove any duplicates and then closes GUI window
		:return: None
		"""
		# Ask user to select target folder
		target_file = tkFileDialog.asksaveasfilename(
			initialdir=self.results_pth,
			defaultextension='.xlsx',
			filetypes=constants.General.file_types,
			title='Please select file for results')

		self.target_file = target_file

		# Process the fault times into useful format converting into floats
		fault_times = self.var_fault_times_list.get()
		fault_times = fault_times.split(',')
		# Loop through each value converting to a float
		# TODO: Add in more error processing here
		self.fault_times = list()
		for val in fault_times:
			try:
				new_val = float(val)
				self.fault_times.append(new_val)
			except ValueError:
				self.logger.warning(
					'Unable to convert the fault time <{}> to a number and so has been skipped'.format(val)
				)

		self.logger.info(
			(
				'Faults will be applied at the busbars listed below and results saved to:\n{} \n '
				'for the fault times: {} seconds.  \nBusbars = \n{}'
			).format(self.target_file, self.fault_times, self.selected_busbars)
		)

		# Destroy GUI
		self.master.destroy()
		return None


class CreateToolTip(object):
	"""
		Function to create a popup tool tip for a given widget based on the descriptions provided here:
			https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter
	"""
	def __init__(self, widget, text='widget info'):
		"""
			Establish link with tooltip
		:param widget:  Tkinter elemnt that tooltip should be associated with
		:param str text:  Message to display when hovering over button
		"""
		self.wait_time = 500     #miliseconds
		self.wrap_length = 450   #pixels
		self.widget = widget
		self.text = text
		self.widget.bind("<Enter>", self.enter)
		self.widget.bind("<Leave>", self.leave)
		self.widget.bind("<ButtonPress>", self.leave)
		self.id = None
		self.tw = None

	def enter(self, event=None):
		del event
		self.schedule()

	def leave(self, event=None):
		del event
		self.unschedule()
		self.hidetip()

	def schedule(self, event=None):
		del event
		self.unschedule()
		self.id = self.widget.after(self.wait_time, self.showtip)

	def unschedule(self, event=None):
		del event
		_id = self.id
		self.id = None
		if _id:
			self.widget.after_cancel(_id)

	def showtip(self):
		x, y, cx, cy = self.widget.bbox("insert")
		x += self.widget.winfo_rootx() + 25
		y += self.widget.winfo_rooty() + 20
		# creates a top level window
		self.tw = tk.Toplevel(self.widget)
		# Leaves only the label and removes the app window
		self.tw.wm_overrideredirect(True)
		self.tw.wm_geometry("+%d+%d" % (x, y))
		label = tk.Label(
			self.tw, text=self.text, justify='left', background="#ffffff", relief='solid', borderwidth=1,
			wraplength = self.wrap_length
		)
		label.pack(ipadx=1)

	def hidetip(self):
		tw = self.tw
		self.tw = None
		if tw:
			tw.destroy()
