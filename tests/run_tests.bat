@ECHO OFF
cd %~dp0%
del output.log
.\scriptunit.exe /Q /log test_results.xml .
type output.log
type test_results.xml

if not "%1%"=="nopause" pause


