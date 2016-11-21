rem uninstall existing service
sc.exe \\krr-app-palbp01 Stop "ArcelorMittal.PrintService"
sc.exe \\krr-app-palbp01 Delete "ArcelorMittal.PrintService"
rem copy new version
xcopy %WORKSPACE%\PrintLabelService\bin\Production\*.* \\krr-app-palbp01\Nikama\print_service /Y
rem install existing service
rem echo off
sc.exe \\krr-app-palbp01 Create "ArcelorMittal.PrintService" binPath="C:\Nikama\print_service\PrintLabelService.exe" start=delayed-auto obj=%PRINT_USER% password=%PRINT_PASS%
sc.exe \\krr-app-palbp01 Start "ArcelorMittal.PrintService"