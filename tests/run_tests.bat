@ECHO OFF
cd %~dp0%
if exist output.log del output.log
.\scriptunit.exe /Q /log test_results.xml .
if exist output.log type output.log
type test_results.xml

if not "%1%"=="nopause" pause


