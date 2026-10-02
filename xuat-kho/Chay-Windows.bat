@echo off
setlocal
chcp 65001 >nul
set "XUAT_KHO_EXIT=1"
cd /d "%~dp0"
if errorlevel 1 goto bad_folder
echo Dang kiem tra Python...
echo Thu vien Excel di kem bo cai; khong can pip hay Internet de khoi dong.
set "XUAT_KHO_PYTHON="
for /f "delims=" %%P in ('py -3 -X utf8 -c "import sys; assert sys.version_info.__ge__((3,10)); print(sys.executable)" 2^>nul') do set "XUAT_KHO_PYTHON=%%P"
if defined XUAT_KHO_PYTHON goto start
for /f "delims=" %%P in ('python -X utf8 -c "import sys; assert sys.version_info.__ge__((3,10)); print(sys.executable)" 2^>nul') do set "XUAT_KHO_PYTHON=%%P"
if defined XUAT_KHO_PYTHON goto start
for /f "delims=" %%P in ('python3 -X utf8 -c "import sys; assert sys.version_info.__ge__((3,10)); print(sys.executable)" 2^>nul') do set "XUAT_KHO_PYTHON=%%P"
if defined XUAT_KHO_PYTHON goto start
echo Khong tim thay Python chay duoc. Can cai Python 3.10 tro len.
echo Khi cai Python, chon Add python.exe to PATH.
echo Tai Python tai https://www.python.org/downloads/
goto finish
:start
"%XUAT_KHO_PYTHON%" -X utf8 "%~dp0khoidong.py"
set "XUAT_KHO_EXIT=%ERRORLEVEL%"
if not "%XUAT_KHO_EXIT%"=="0" echo Neu co loi, gui noi dung cua so nay hoac file Loi-khoi-dong.log.
goto finish
:bad_folder
echo Khong mo duoc thu muc ung dung. Hay giai nen toan bo file ZIP vao thu muc moi.
:finish
echo.
echo Nhan phim bat ky de dong cua so nay.
pause >nul
endlocal & exit /b %XUAT_KHO_EXIT%
