:BEGIN
CLS
set hh=%time:~0,2%
ECHO %hh%
SET apppath=%~dp0
ECHO app path = %apppath%
cd %apppath%
SET venv_path=%apppath%\venv
python -m venv %venv_path%
CALL %venv_path%\Scripts\activate.bat
CALL python -m pip install --upgrade pip
CALL pip install -r requirements.txt

:EXIT



