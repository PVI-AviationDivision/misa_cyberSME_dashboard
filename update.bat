@echo off
setlocal
cd /d "%~dp0"
title An Tam So - Cap nhat dashboard
echo.
echo  =====================================================
echo    BH PVI x MISA  ^|  Bao hiem An Tam So - Daily Ops
echo  =====================================================
echo.

if not exist "inbox\*.xlsx" (
  echo  [!] Thu muc inbox dang trong - se dung lai du lieu da luu trong archive.
  echo.
)

echo  [1/3] Doc bao cao Excel va tao du lieu dashboard...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build.ps1"
if errorlevel 1 (
  echo.
  echo  [X] Loi khi doc bao cao. Kiem tra lai file trong inbox.
  pause
  exit /b 1
)

echo.
echo  [2/3] Day len GitHub...
where git >nul 2>&1
if errorlevel 1 (
  echo  [!] May chua cai Git - chi cap nhat ban xem tren may. Cai Git: https://git-scm.com/download/win
  goto open
)
if not exist ".git" (
  echo  [i] Lan dau chay: ket noi thu muc voi repo GitHub...
  git init -q
  git symbolic-ref HEAD refs/heads/main
  git remote add origin https://github.com/PVI-AviationDivision/misa_cyberSME_dashboard.git
  git fetch -q origin 2>nul
  git rev-parse -q --verify origin/main >nul 2>&1 && git reset -q origin/main
)
set GITMAIL=
for /f "delims=" %%e in ('git config user.email') do set GITMAIL=%%e
if not defined GITMAIL (
  git config user.name "BH PVI - P.TMDT"
  git config user.email "dashboard@users.noreply.github.com"
)
git add -A
git commit -q -m "Cap nhat so lieu %date% %time:~0,5%"
git pull -q --rebase origin main 2>nul
git push -q -u origin HEAD:main
if errorlevel 1 (
  echo  [X] Push len GitHub that bai - kiem tra mang / dang nhap GitHub roi chay lai.
) else (
  echo  [OK] Da day len GitHub. Trang online cap nhat sau khoang 1-2 phut.
)

:open
echo.
echo  [3/3] Mo dashboard - bam nut "Tao tin nhan Zalo" de lay mau tin nhan.
start "" "%~dp0index.html"
echo.
echo  Link bao cao: https://pvi-aviationdivision.github.io/misa_cyberSME_dashboard/
echo.
pause
