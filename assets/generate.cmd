setlocal enabledelayedexpansion
@echo off

:: icon text
set name=icon

set THISDIR=%~dp0
set THISDIR=%THISDIR:~,-1%

:: Path to inkscape install
set inkscape="%tools%\Programs\inkscape\inkscape.exe"

for %%s in (16 32 64 80 128 300) do (
    set size=%%s
    set command=%inkscape% -z "%THISDIR%/%name%.svg" -w !size! -h !size! -e "%THISDIR%/%name%-!size!.png"
    echo !command!
    call !command!
)
