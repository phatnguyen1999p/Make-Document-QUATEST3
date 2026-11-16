@echo off
rem Khoi dong tien trinh nen (an cua so) de Claude tu chay ho so, khong can dung vao File Explorer.
rem Chi can bam dup 1 lan sau khi bat may. Dung lai: stop_lamkq_watcher.bat
cd /d "%~dp0"
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0lamkq_watcher.ps1"
echo Da khoi dong lamkq_watcher (chay nen, an). Co the dong cua so nay.
timeout /t 3 >nul
