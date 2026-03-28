@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

set "ENTRY=src\gui_app.py"
set "EXE_ASCII=CommentAggregator"

echo.
echo ====================================================
echo  Comment Sheet Aggregator - Build Script
echo ====================================================
echo.

echo [0/3] Checking PyInstaller...
pyinstaller --version > nul 2>&1
if %errorlevel% neq 0 (
    echo       Not found. Installing...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install PyInstaller.
        goto :FAIL
    )
)
echo       OK
echo.

echo [1/3] Removing previous build artifacts...
if exist "build"             rmdir /s /q "build"
if exist "dist"              rmdir /s /q "dist"
if exist "%EXE_ASCII%.spec"  del   /q    "%EXE_ASCII%.spec"
echo       OK
echo.

echo [2/3] Building EXE...
echo.

pyinstaller --onefile --noconsole --clean ^
    --name "%EXE_ASCII%" ^
    --paths "src" ^
    --add-data "xlrd_legacy;xlrd_legacy" ^
    --hidden-import=aggregator ^
    --hidden-import=xlrd ^
    --hidden-import=openpyxl ^
    --hidden-import=pandas ^
    "%ENTRY%"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed.
    goto :CLEANUP_AND_FAIL
)
echo.

echo [3/3] Moving EXE to root and cleaning up...

move /Y "dist\%EXE_ASCII%.exe" "%EXE_ASCII%.exe" > nul
if %errorlevel% neq 0 (
    echo [ERROR] Failed to move EXE.
    goto :CLEANUP_AND_FAIL
)

python -c "import os; os.replace('%EXE_ASCII%.exe', '\u30b3\u30e1\u30f3\u30c8\u30b7\u30fc\u30c8\u96c6\u8a08\u30c4\u30fc\u30eb.exe')"

if exist "build"             rmdir /s /q "build"
if exist "dist"              rmdir /s /q "dist"
if exist "%EXE_ASCII%.spec"  del   /q    "%EXE_ASCII%.spec"

for /d /r "src" %%D in (__pycache__) do (
    if exist "%%D" rmdir /s /q "%%D"
)

echo       OK
echo.
echo ====================================================
echo  Build complete!
echo  Output: %~dp0
echo ====================================================
echo.
pause
exit /b 0


:CLEANUP_AND_FAIL
if exist "build"             rmdir /s /q "build"
if exist "dist"              rmdir /s /q "dist"
if exist "%EXE_ASCII%.spec"  del   /q    "%EXE_ASCII%.spec"

:FAIL
echo.
echo ====================================================
echo  Build FAILED. See error messages above.
echo ====================================================
echo.
pause
exit /b 1
