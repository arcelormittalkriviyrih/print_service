rem uninstall existing service
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /u "C:\Nikama\print_service\PrintLabelService.exe"
rem copy new version
xcopy %WORKSPACE%\PrintLabelService\bin\Debug\*.* C:\Nikama\print_service /Y
rem install existing service
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /username=europe\krr-svc-palbp-auto /password=%password% /unattended "C:\Nikama\print_service\PrintLabelService.exe"
rem start service
net start "ArcelorMittal.PrintService"