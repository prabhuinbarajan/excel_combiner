@ECHO OFF

:BEGIN
CLS
set "app_env=%1"
if "%app_env%"=="" set "app_env=PROD_TEST"
set "app_env=PROD_TEST"
echo %app_env%
set hh=%time:~0,2%
ECHO %hh%
SET apppath=%~dp0
ECHO app path = %apppath%
cd %apppath%
SET venv_path=%apppath%\venv
if "%time:~0,1%"==" " set hh=0%hh:~1,1%
> "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\logs\consol_log-%app_env%-%DATE:~-4%-%DATE:~4,2%-%DATE:~7,2%_%hh%%time:~3,2%%time:~6,2%.txt" 2>&1 (

CALL %venv_path%\Scripts\activate.bat
CALL pip install -r requirements.txt
ECHO *** MAIN SCRIPT START  ***
ECHO %date:~10,4%%date:~4,2%%date:~7,2%_%time:~0,2%%time:~3,2%%time:~6,2%

ECHO *** CALLING POST PROCESSING SCRIPTS ***
ECHO *** CALLING MULTIPLE BU PROCESSING SCRIPT ***
python test1.py %app_env%
python tb_multiple_bus_lc.py %app_env%
REM python tb_a510_bs_trend.py %app_env%
REM python intl_vol_rate.py %app_env%
REM python comp_and_ben_detail.py %app_env%
REM python comm_fee_rate_vol.py %app_env%
REM python a510_bs_detail_seg.py %app_env%
ECHO *** POST PROCESSING SCRIPTS COMPLETE***
ECHO *** CALLING FOLDER RENAME SCRIPT ***
REM python folder_rename.py %app_env%
ECHO *** FOLDER RENAME SCRIPT COMPLETE***
deactivate
ECHO *** SCRIPT COMPLETED  ***
)
:EXIT
          