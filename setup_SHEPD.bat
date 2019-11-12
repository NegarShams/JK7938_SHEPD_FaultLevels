REM
set current_dir=%cd%
set target_dir=%current_dir%\local_packages

if not exist %target_dir% mkdir %target_dir%

set pip_execute=%current_dir%\pip-19.3.1-py2.py3-none-any.whl/pip
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\setuptools-41.6.0-py2.py3-none-anypip.whl
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\six-1.12.0-py2.py3-none-any.whl
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\numpy-1.16.5-cp27-cp27m-win32.whl
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\xlrd-1.2.0-py2.py3-none-any.whl
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\xlwt-1.3.0-py2.py3-none-any.whl
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\python_dateutil-2.8.1-py2.py3-none-any.whl
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\pytz-2019.3-py2.py3-none-any.whl
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\et_xmlfile-1.0.1.tar.gz
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\openpyxl-2.6.4.tar.gz
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\jdcal-1.4.1-py2.py3-none-any.whl
python %pip_execute% install --no-deps --target=%target_dir% --upgrade --force-reinstall %current_dir%\pandas-0.24.2-cp27-cp27m-win32.whl

