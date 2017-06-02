rem uninstall existing service
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /u "C:\Nikama\print_service\PrintLabelService.exe"
rem copy new version
xcopy %WORKSPACE%\PrintLabelService\bin\Release\*.* C:\Nikama\print_service /Y
rem install existing service
echo off
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /username=%ADMIN_USER% /password=%ADMIN_PASS% /unattended "C:\Nikama\print_service\PrintLabelService.exe"
echo on
rem first run with administrator privileges
net start "ArcelorMittal.PrintService"
net stop "ArcelorMittal.PrintService"
sc.exe config "ArcelorMittal.PrintService" obj=%PRINT_USER% password=%PRINT_PASS%
rem configure delayed service
rem sc.exe config "ArcelorMittal.PrintService" start=delayed-auto
sc.exe config "ArcelorMittal.PrintService" start=demand
rem net stop "ArcelorMittal.PrintService"
