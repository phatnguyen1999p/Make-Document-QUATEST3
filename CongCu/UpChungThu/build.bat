@echo off
chcp 65001 > nul
echo ========================================================
echo   QUATEST 3 - HỆ THỐNG XỬ LÝ VĂN BẢN - ver: 6.8.1
echo              TIẾN HÀNH COMPILE PHẦN MỀM
echo ========================================================
echo.

echo [1/3] Kiểm tra cài đặt PyInstaller...
python -m pip show pyinstaller > nul 2>&1
if %errorlevel% neq 0 (
    echo Chưa cài đặt PyInstaller. Đang tiến hành cài đặt...
    python -m pip install pyinstaller
)

echo.
echo [2/3] Đang tiến hành build phần mềm bằng PyInstaller...
pyinstaller --noconfirm UpChungThu_v6.spec

if %errorlevel% equ 0 (
    echo.
    echo ========================================================
    echo  [3/3] COMPILE THÀNH CÔNG!
    echo  File thực thi nằm tại: dist\UpChungThu_v6\UpChungThu_v6.exe
    echo ========================================================
) else (
    echo.
    echo ========================================================
    echo  [!] COMPILE THẤT BẠI! Vui lòng kiểm tra lại lỗi trên.
    echo ========================================================
)

pause
