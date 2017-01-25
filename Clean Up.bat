@echo off
del /f /s *.tmp
del /f /s *.log

set curdir=%~dp0
cd /d %curdir%\Components
del /f 工程1.vbp
del /f 工程1.vbw
del /f Form1.frm
del /f Form1.frx
del /f Module1.bas