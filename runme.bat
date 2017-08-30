@echo off
@setlocal enableextensions
@cd /d "%~dp0
set script_path=%~dp0

echo ============================================================================
echo  TIMLAB System Security Vision (SSV)
echo  timip.net
echo  By Tim Ip
echo =============================================================================
echo.
echo  [+] Start Time = %date% %time%
echo  [+] Machine Name = %computername%
echo  [+] Script Path = %script_path%

for %%f in (*.vbs) do (
	set /p val=<%%f
	echo  [+] Running %%f...
	cscript //nologo %script_path%%%f > %%~nf_%COMPUTERNAME%.ssv 2>&1
)

echo  [+] Done.
echo  [+] Finish Time = %date% %time%
