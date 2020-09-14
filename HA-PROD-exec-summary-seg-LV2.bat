@ECHO OFF

:BEGIN
CLS
set "app_env=%1"
if "%app_env%"=="" set "app_env=PROD_TEST"
REM set "app_env=PROD_TEST"
echo %app_env%
set hh=%time:~0,2%
ECHO %hh%
SET apppath=%~dp0
ECHO app path = %apppath%
cd %apppath%
SET venv_path=%apppath%\venv
if "%time:~0,1%"==" " set hh=0%hh:~1,1%
> "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\logs\EXEC_Summary_Seg_consol_log-%app_env%-%DATE:~-4%-%DATE:~4,2%-%DATE:~7,2%_%hh%%time:~3,2%%time:~6,2%.txt" 2>&1 (

CALL %venv_path%\Scripts\activate.bat
CALL pip install -r requirements.txt
cmd /v /c "echo !date!:!time!"
ECHO *** CALLING Exec Summary Segment LV2 SCRIPT ***
python exec_summary_segment_LV2.py %app_env%
ECHO *** Exec Summary Segment LV2 SCRIPT COMPLETE***
cmd /v /c "echo !date!:!time!"
deactivate
ECHO *** MAIN SCRIPT COMPLETED  ***
)
:EXIT

          