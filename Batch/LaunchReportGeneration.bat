@echo off

REM ****************************************************************
REM **** PRE-REQUISITES 										****
REM **** 7Zip installation : https://7-zip.org/download.html	****
REM **** Microsoft Excel											****
REM **** At least 2 snapshots 									****
REM **** At least 1 rule in the Education Plan					****
REM ****************************************************************

REM ****************************************************************
REM **** PARAMETERS	 											****
REM ****  #1 : 	the full path & name of the configuration file	****
REM **** Example:	C:\Report\MyConfig.txt						****
REM ****************************************************************

REM ****************************************************************
REM **** REMINDER	BE CAREFUL	DON'T FORGET					****
REM **** Before launching the batch program, please set up 		****
REM **** the proper execution parameter in the configuration 	****
rem **** file passed as a parameter								****
REM ****************************************************************

if "%~1"=="" (
	echo Missing Parameter
	echo The configuration file name and path must be passed as a parameter to the LaunchReportGeneration.bat program
	echo Example:	LaunchReportGeneration.bat C:\Report\MyConfig.txt
	goto AutomationFailed )

REM Set the program variables defined in the config file
for /f "delims=" %%x in (%1) do (set "%%x")

for /f %%i in ('WMIC OS GET LocalDateTime ^| FIND "."') DO SET CaptureDate=%%i
set CaptureDate=%CaptureDate:~0,4%%CaptureDate:~4,2%%CaptureDate:~6,2%%CaptureDate:~8,2%%CaptureDate:~10,2%

if not exist "%ReportLogPath%" mkdir "%ReportLogPath%"
REM Ping to kill some time
@ping 127.0.0.1 -n 2 -w 10000 > nul
if not exist "%ReportPath%" mkdir "%ReportPath%"	> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log

echo Generating the CSV Reports
echo Generating the CSV Reports  >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log

call %BatchPath%\GenerateCSVReports.bat %ReportPath% %BatchPath% %DashboardService% %ApplicationName% %CaptureDate% %CSSServer% %userLogin% %database% %portNumber% %EDURL% %CCSPSQL% >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log
if %errorlevel% NEQ 0 echo FAILED CSV Generation & goto :AutomationFailed

echo SUCCESSFUL CSV Generation
echo SUCCESSFUL CSV Generation >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log

echo Zipping the CSV reports >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log
%SevenZip% a -tzip "%ReportPath%\%ApplicationName%_NewViolationsReport_%CaptureDate%.zip" "%ReportPath%\%ApplicationName%_*_%CaptureDate%.csv" >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log
if %errorlevel% NEQ 0 goto :AutomationFailed

echo SUCCESSFUL CSV Zip >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log

:LaunchVBScript

REM ApplicationName,CaptureDate,ReportPath,ReportLogPath
echo Generating Excel Report on %ApplicationName% in folder %ReportPath%
echo Generating Excel Report on %ApplicationName% in folder %ReportPath%  >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log
cscript %BatchPath%\GenerateExcelReport.vbs "%ApplicationName%" %CaptureDate% "%BatchPath%" "%ReportLogPath%" "%ReportPath%"
if %errorlevel% NEQ 0 goto :AutomationFailed

echo SUCCESSFUL Excel Report Generation
echo SUCCESSFUL Excel Report Generation >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log

echo Removing the CSV Reports >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log
del "%ReportPath%\%ApplicationName%_*_%CaptureDate%.csv" >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log
echo SUCCESSFUL CSV Reports Removal >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log

goto :end

REM echo SENDING THE REPORTS to %ReportMail%
REM C:\Customers\CheckStyle\Automation\Email\postie -host:"EXCH1.corp.castsoftware.com" -from:"jenkins@bdrocks2.com" -to:"%ReportMail%" -s:"%ApplicationName%: Violations Report from %CaptureDate%" -msg:"Hi, Attached are the %ApplicationName% violation reports done on %CaptureDate%." -a:"%ReportPath%\%ApplicationName%_ViolationsReport_%CaptureDate%.zip"
REM if %errorlevel% NEQ 0 goto :AutomationFailed


REM echo SENDING SUCCESSFUL STATUS to %StatustMail%
REM C:\Customers\CheckStyle\Automation\Email\postie -host:"EXCH1.corp.castsoftware.com" -from:"jenkins@bdrocks2.com" -to:"%StatusMail%" -s:"%ApplicationName% - %CaptureDate%: Analysis Automation Successful" -msg:"Hi, the %ApplicationName% analysis was successful and the reports sent to %ReportMail%." -a:"%ReportPath%\%ApplicationName%_ViolationsReport_%CaptureDate%.zip"
REM :AutomationSuccessful

:AutomationFailed
REM echo SENDING FAILED STATUS to %StatustMail%
REM C:\Customers\CheckStyle\Automation\Email\postie -host:"EXCH1.corp.castsoftware.com" -from:"jenkins@bdrocks2.com" -to:"%StatusMail%" -s:"%ApplicationName% - %CaptureDate%: Analysis Automation Failed" REM -msg:"Hi, the %ApplicationName% analysis has failed." 

echo FAILED Report Generation
echo FAILED Report Generation >> %ReportLogPath%\NewViolationsReport_%ApplicationName%_%CaptureDate%.log
exit /B 1

:end
exit /B 0
