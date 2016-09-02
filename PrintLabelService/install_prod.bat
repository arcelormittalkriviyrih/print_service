rem uninstall existing service
sc.exe \\krr-app-palbp01 Delete "ArcelorMittal.PrintService_test"
rem copy new version
xcopy %WORKSPACE%\PrintLabelService\bin\Production\*.* \\krr-app-palbp01\Nikama\print_service_test /Y
rem install existing service
rem echo off
sc.exe \\krr-app-palbp01 Create "ArcelorMittal.PrintService_test" binPath="C:\Nikama\print_service_test\PrintLabelService.exe" start=demand obj=%PRINT_USER% password=%PRINT_PASS%
