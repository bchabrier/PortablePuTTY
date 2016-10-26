set "message=%1%"
set "buttons=%2%"
set "title=%3%"

set "vbOK=1"
set "vbCancel=2"
set "vbAbort=3"
set "vbRetry=4"
set "vbIgnore=5"
set "vbYes=6"
set "vbNo=7"

if %title%=="Use local sessions?" exit /b %vbYes%

exit /b 9999

