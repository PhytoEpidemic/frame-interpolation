@echo off

set DIR=%~dp0envr

set PATH=C:\Windows\system32;C:\Windows;%DIR%\frame_interpolation;%DIR%\frame_interpolation\Scripts;%DIR%\frame_interpolation\Scripts\Library\bin
set PY_LIBS=%DIR%\frame_interpolation\Scripts\Lib;%DIR%\frame_interpolation\Scripts\Lib\site-packages
set PY_PIP=%DIR%\frame_interpolation\Scripts
::set PIP_INSTALLER_LOCATION=%DIR%\python\get-pip.py
