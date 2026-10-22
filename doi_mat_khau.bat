@echo off
setlocal
cd /d "%~dp0"
title An Tam So - Doi mat khau dashboard
echo.
echo  Doi mat khau xem dashboard online.
echo  Sau khi doi, moi nguoi (ke ca tren iPhone) phai nhap mat khau moi.
echo.
attrib -h ".secret" >nul 2>&1
del ".secret" >nul 2>&1
call "%~dp0update.bat"
