@ECHO OFF

:BEGIN
CLS
set hh=%time:~0,2%
ECHO %hh%
if "%time:~0,1%"==" " set hh=0%hh:~1,1%
> "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\logs\consol_log-%DATE:~-4%-%DATE:~4,2%-%DATE:~7,2%_%hh%%time:~3,2%%time:~6,2%.txt" 2>&1 (
ECHO *** MAIN SCRIPT START  ***
ECHO %date:~10,4%%date:~4,2%%date:~7,2%_%time:~0,2%%time:~3,2%%time:~6,2%

ECHO *** CALLING POST PROCESSING SCRIPTS ***
ECHO *** CALLING MULTIPLE BU PROCESSING SCRIPT ***
CALL "C:\projects\excelcombiner\Scripts\python.exe" "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\PostProcessingScripts\tb_multiple_bus_lc.py" PROD_TEST
REM ECHO *** CALLING International Volume Rate SCRIPT ***
REM "C:\projects\excelcombiner\Scripts\python.exe" "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\PostProcessingScripts\intl_vol_rate_v2.py" %ENV% %MYPER% %MYYEAR%
REM "C:\projects\excelcombiner\Scripts\python.exe" "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\PostProcessingScripts\tb_multiple_bus_lc_v2.py" %ENV% %MYPER% %MYYEAR%
REM "C:\projects\excelcombiner\Scripts\python.exe" "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\PostProcessingScripts\tb_multiple_bus_lc_v2.py" %ENV% %MYPER% %MYYEAR%
REM "C:\projects\excelcombiner\Scripts\python.exe" "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\PostProcessingScripts\tb_multiple_bus_lc_v2.py" %ENV% %MYPER% %MYYEAR%
REM "C:\projects\excelcombiner\Scripts\python.exe" "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\PostProcessingScripts\tb_multiple_bus_lc_v2.py" %ENV% %MYPER% %MYYEAR%
REM "C:\projects\excelcombiner\Scripts\python.exe" "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\PostProcessingScripts\tb_multiple_bus_lc_v2.py" %ENV% %MYPER% %MYYEAR%
ECHO *** POST PROCESSING SCRIPTS COMPLETE***
ECHO *** CALLING FOLDER RENAME SCRIPT ***
CALL "C:\projects\excelcombiner\Scripts\python.exe" "\\sclbonas01\DATA\ERP\Accounting\General Ledger\Period End Close\Host Analytics\PostProcessingScripts\folder_rename.py" PROD_TEST
ECHO *** FOLDER RENAME SCRIPT COMPLETE***
ECHO *** SCRIPT COMPLETED  ***
)
:EXIT
          