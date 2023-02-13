@echo off
setlocal
set THISDIR=%~dp0
set THISDIR=%THISDIR:~,-1%

call %scripts%\md-to-html.cmd %THISDIR%\privacy.md --css style.css --metadata=pagetitle:"Privacy"
call %scripts%\md-to-html.cmd %THISDIR%\support.md --css style.css --metadata=pagetitle:"Support"
call %scripts%\md-to-html.cmd %THISDIR%\eula.md --css style.css --metadata=pagetitle:"EULA"