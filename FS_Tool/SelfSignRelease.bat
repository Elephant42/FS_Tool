@echo off

SET CERT_PATH=Cert:\CurrentUser\My\*99E9FCA2F8D1EFBD1BFB5A62D85CC1E1DB3438FA
SET TS_SVR=http://timestamp.digicert.com

SET TARGET_DIR=bin\Release
xcopy LICENSE "%TARGET_DIR%" /y
xcopy ..\README.md "%TARGET_DIR%" /y
powershell CodeSign.ps1 -file_path "%TARGET_DIR%\FS_Tool.exe" -cert_path "%CERT_PATH%" -ts_svr "%TS_SVR%"
powershell CodeSign.ps1 -file_path "%TARGET_DIR%\SimConnectLib.dll" -cert_path "%CERT_PATH%" -ts_svr "%TS_SVR%"
powershell CodeSign.ps1 -file_path "%TARGET_DIR%\HidSharp.dll" -cert_path "%CERT_PATH%" -ts_svr "%TS_SVR%"

pause
