rem uninstall existing service
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /u "C:\Nikama\print_service\PrintLabelService.exe"
rem copy new version
xcopy %WORKSPACE%\PrintLabelService\bin\Release\*.* C:\Nikama\print_service /Y
rem install existing service
echo off
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /username=%PRINT_USER% /password=%PRINT_PASS% /unattended "C:\Nikama\print_service\PrintLabelService.exe"
echo on
net stop "ArcelorMittal.PrintService"