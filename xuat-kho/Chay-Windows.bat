@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
echo Dang kiem tra Python...
py -3 -c "import sys; assert sys.version_info >= (3,10)" >nul 2>nul
if not errorlevel 1 goto use_py
python -c "import sys; assert sys.version_info >= (3,10)" >nul 2>nul
if not errorlevel 1 goto use_python
python3 -c "import sys; assert sys.version_info >= (3,10)" >nul 2>nul
if not errorlevel 1 goto use_python3
echo Khong tim thay Python chay duoc. Can cai Python 3.10 tro len.
echo Khi cai Python, chon Add python.exe to PATH.
goto finish
:use_py
py -3 -c "import khoidong; raise SystemExit(khoidong.main())"
goto finish
:use_python
python -c "import khoidong; raise SystemExit(khoidong.main())"
goto finish
:use_python3
python3 -c "import khoidong; raise SystemExit(khoidong.main())"
:finish
echo.
echo Ung dung da dung. Neu co loi, gui noi dung trong cua so nay.
pause
endlocal
