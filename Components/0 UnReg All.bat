@echo off
for /f "delims=" %%i in ('dir ..\Dlls\*.dll /b /s') do regsvr32 /s /u "%%i"&&echo.成功反注册 %%i
for /f "delims=" %%i in ('dir ..\Dlls\*.ocx /b /s') do regsvr32 /s /u "%%i"&&echo.成功反注册 %%i