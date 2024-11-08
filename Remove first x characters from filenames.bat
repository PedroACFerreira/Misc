for %%i in ("*.mp3") do (set fname=%%i) & call :rename
goto :eof
:rename
::Cuts off 1st four chars, then appends prefix
ren "%fname%" "%fname:~4%"
goto :eof