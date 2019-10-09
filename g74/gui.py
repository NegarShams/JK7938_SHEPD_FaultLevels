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
import sys
import os
import logging
import g74.constants as constants

logger = logging.getLogger()


class MainGUI:
	"""
		Main class to produce the GUI
		Allows the user to select the busbars and methodology to be applied in the fault current calculations
	"""
	def __init__(self, title=constants.GUI.gui_name):
		"""
			Initialise GUI
		:param str title: (optional) - Title to be used for main window
		"""
		# General constants which need to be initialised
		self._row = 0
		self._col = 0

		# Is populated with a list of file paths to be returned
		self.selected_busbars = list()
		# Target file to export results to
		self.target_file = str()

		# Initialise constants and tk window
		self.master = tk.Tk()
		self.master.title = title

		self.cmd_add_file = tk.Button(self.master,
										   text=lbl_button,
										   command=self.add_new_file)
		self.cmd_add_file.grid(row=self.row(), column=self.col())

		_ = tk.Label(master=self.master,
						  text='Files to be compared:')
		_.grid(row=self.row(1), column=self.col())

		self.lbl_results_files = tkinter.scrolledtext.ScrolledText(master=self.master)
		self.lbl_results_files.grid(row=self.row(1), column=self.col())
		self.lbl_results_files.insert(tk.INSERT,
									  'No Folders Selected')

		self.cmd_process_results = tk.Button(self.master,
												  text='Import',
												  command=self.process)
		self.cmd_process_results.grid(row=self.row(1), column=self.col())

		logger.debug('GUI window created')
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

	def add_new_file(self):
		"""
			Function to load Tkinter.askopenfilename for the user to select a file and then adds to the
			self.file_list scrolling text box
		:return: None
		"""
		# Ask user to select file(s) or folders based on <.bo_files>
		if self.bo_files:
			# User will be able to select list of files
			file_paths = file_selector(initial_pth=self.results_pth, open_file=True)
		else:
			# User will be able to add folder
			file_paths = file_selector(initial_pth=self.results_pth, save_dir=True,
									  lbl_folder_select='Select folder containing hast results files')

		# User can select multiple and so following loop will add each one
		for file_pth in file_paths:
			logger.debug('Results folder {} added as input folder'.format(file_pth))
			self.results_pth = os.path.dirname(file_pth)

			# Add complete file pth to results list
			if not self.results_files_list:
				# If initial list is empty then will need to replace with initial string
				self.lbl_results_files.delete(1.0, tk.END)

			self.results_files_list.append(file_pth)
			self.lbl_results_files.insert(tk.END,
										  '{} - {}\n'.format(len(self.results_files_list), file_pth))

	def process(self):
		"""
			Function sorts the files list to remove any duplicates and then closes GUI window
		:return: None
		"""
		# Sort results into a single list and remove any duplicates
		self.results_files_list = list(set(self.results_files_list))

		# Ask user to select target folder
		target_file = file_selector(initial_pth=self.results_pth, save_file=True,
									lbl_file_select='Please select file for results')[0]

		# Check if user has input an extension and if not then add it
		file, _ = os.path.splitext(target_file)

		self.target_file = file + self.ext

		# Destroy GUI
		self.master.destroy()
