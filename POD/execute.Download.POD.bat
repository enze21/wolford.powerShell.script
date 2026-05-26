
@echo off
set LOGFILE=log_esecuzione.txt

echo Esecuzione iniziata: %date% %time% > %LOGFILE%
powershell -ExecutionPolicy Bypass -File "D:\POD\Script\downLoad.POD.files.ps1" >> %LOGFILE% 2>&1
echo Esecuzione terminata: %date% %time% >> %LOGFILE%

:: Ottieni la data corrente in formato YYYYMMDD
for /f %%i in ('powershell -command "Get-Date -Format yyyyMMdd"') do set "DATA=%%i"

:: Costruisci il nuovo nome
set "NUOVO=log_esecuzione_%DATA%.txt"

:: Rinomina il file
rename "%LOGFILE%" "%NUOVO%"

echo File rinominato in: %NUOVO%
