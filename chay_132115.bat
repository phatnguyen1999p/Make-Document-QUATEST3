@echo off
rem Bam dup file nay de chay ho so 132115. Khong dong bo o mang Z:.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0chay_132115.ps1" %*
echo.
echo Xong. Xem chay_132115_log.txt
pause
