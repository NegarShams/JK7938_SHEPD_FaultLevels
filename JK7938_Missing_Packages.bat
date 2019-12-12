REM
REM ###################################################################################################################
REM ###										PSSE G74 Fault Studies													###
REM ###		Script sets up Python on SHEPD machines for use with PSSE and G74 script produced by PSC.               ###
REM ###		This is necessary to make available the following packages to the G74 script:							###
REM ###			- setuptools																						###
REM ###			- six																								###
REM ###			- numpy																								###
REM ###			- xlrd																								###
REM ###			- xlwt																								###
REM ###			- dateutil																							###
REM ###			- pytz  																							###
REM ###			- et_emlfile																						###
REM ###			- openpyxl  																						###
REM ###			- jdcal       																						###
REM ###			- pandas       																						###
REM ###																												###
REM ###		These packages will be installed in the same directory as this batch file is run and so should not      ###
REM ###     administrative privaledges if the batch file is located in the users directory. The package 			###
REM ###     installation is completed using python 2.7 and pip, since sometimes the SHEPD installation does not 	###
REM ###     have pip installed it is provided as part of this package.                                           	###
REM ###																												###
REM ###		Also, since python is not defined as an environment variable on the SHEPD installation this batch file  ###
REM ###		will search for python 2.7 installation on the C drive.  If a different drive is necessary then the     ###
REM ###		user will need to change the search path                                                                ###
REM ###																												###
REM ###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 	###
REM ###		project JK7938 - SHEPD - studies and automation															###
REM ###																												###
REM ###################################################################################################################
REM """

@echo off
echo "Defining target directories for python package installation and finding python executable"
REM define variables for current directory and target directory for the packages to be installed in
set current_dir="%cd%"
set target_dir=%current_dir%\local_packages

echo "Python packages will be installed here: "%target_dir%

REM If the target directory for the new scripts does not exist then this will be created
if not exist %target_dir% mkdir %target_dir%

REM To run the scripts python is required and will now be found by searching the computer with the assumption
REM that a python installation is available on the C drive.
REM Loops through each folder looking for a folder which references Python27 and contains a python.exe executable
for /d /r "c:\" %%a in (*) do (
    if /i "%%~nxa"=="Python27" (
        set folderpath=%%a
        call :innerloop
        )
    )

REM Inner loop used so that a goto break command is possible.
REM Just defines the path to the python file that will be used for the installation
:innerloop
if exist %folderpath%\python.exe (
    set pythonpath=%folderpath%\python.exe
    goto :break
    )

REM Break command to exit for loop once python file has been found to avoid continuing to search
:break
echo "Python installation found here: ""%pythonpath%

REM Define pip executable file which will be used to install all the wheels
set pip_execute=%current_dir%\pip-19.3.1-py2.py3-none-any.whl/pip

REM Install each of the defined wheel files in the local directory, forcing a replacement if they already exist
REM and avoiding the downloading of dependencies.  Therefore all dependencies will need to be installed manually in
REM here.
REM Output turned on so progress of installation can be monitored
echo on
echo "Any errors mean a package did not install and should be investigated further"
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\setuptools-41.6.0-py2.py3-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\six-1.12.0-py2.py3-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\numpy-1.16.5-cp27-cp27m-win32.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\xlrd-1.2.0-py2.py3-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\xlwt-1.3.0-py2.py3-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\python_dateutil-2.8.1-py2.py3-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\pytz-2019.3-py2.py3-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\et_xmlfile-1.0.1-cp27-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\openpyxl-2.6.4-py2.py3-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\jdcal-1.4.1-py2.py3-none-any.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\pandas-0.24.2-cp27-cp27m-win32.whl
%pythonpath% %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\Pillow-6.2.1-cp27-cp27m-win32.whl

echo "All python packages have been installed."
exit
