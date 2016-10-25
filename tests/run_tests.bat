@ECHO OFF
cd %~dp0%
.\scriptunit.exe /Q /log test_results.xml .
type test_results.xml

if not "%1%"=="nopause" pause


