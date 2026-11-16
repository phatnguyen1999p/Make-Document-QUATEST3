@echo off
rem Dung tien trinh nen lamkq_watcher.
cd /d "%~dp0"
echo stop> "%~dp0lamkq_stop.txt"
echo Da gui lenh dung (watcher dung sau toi da vai giay).
timeout /t 3 >nul
