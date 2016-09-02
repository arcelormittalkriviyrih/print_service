rem uninstall existing service
sc.exe \\krr-app-palbp01 Delete "ArcelorMittal.PrintService"
rem copy new version
xcopy %WORKSPACE%\PrintLabelService\bin\Production\*.* \\krr-app-palbp01\Nikama\print_service /Y
rem install existing service
rem echo off
sc.exe \\krr-app-palbp01 Create "ArcelorMittal.PrintService" binPath="C:\Nikama\print_service\PrintLabelService.exe" start=auto obj=%PRINT_USER% password=%PRINT_PASS%
