@echo off
:: ============================================================
:: RemitSaver – COM Add-in Registration
:: Run this script as Administrator.
:: ============================================================

setlocal EnableDelayedExpansion

echo ============================================================
echo  RemitSaver Installer
echo ============================================================
echo.

:: ── Verify we are running as Administrator ────────────────────────────────
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo ERROR: This script must be run as Administrator.
    echo Right-click install.bat and choose "Run as administrator".
    pause
    exit /b 1
)

:: ── Locate Python ────────────────────────────────────────────────────────
set PYTHON=
for %%P in (python python3) do (
    %%P --version >nul 2>&1
    if !errorLevel! == 0 (
        set PYTHON=%%P
        goto :found_python
    )
)
echo ERROR: Python not found on PATH.
echo Install Python 3.9+ and ensure it is on your PATH, then re-run.
pause
exit /b 1
:found_python
echo [OK] Python found: %PYTHON%

:: ── Verify pywin32 ────────────────────────────────────────────────────────
%PYTHON% -c "import win32com" >nul 2>&1
if %errorLevel% NEQ 0 (
    echo [INFO] pywin32 not found – installing dependencies…
    %PYTHON% -m pip install -r "%~dp0requirements.txt"
    if %errorLevel% NEQ 0 (
        echo ERROR: pip install failed.
        pause
        exit /b 1
    )
)
echo [OK] Dependencies verified.

:: ── Run the pywin32 post-install script (needed for COM registration) ──────
set SCRIPTS_DIR=
for /f "delims=" %%S in ('%PYTHON% -c "import sysconfig; print(sysconfig.get_path(\"scripts\"))"') do set SCRIPTS_DIR=%%S
if exist "%SCRIPTS_DIR%\pywin32_postinstall.py" (
    echo [INFO] Running pywin32 post-install…
    %PYTHON% "%SCRIPTS_DIR%\pywin32_postinstall.py" -install >nul 2>&1
)

:: ── Register the COM add-in ────────────────────────────────────────────────
echo [INFO] Registering RemitSaver COM add-in…
%PYTHON% "%~dp0remitsaver_addin.py" --register
if %errorLevel% NEQ 0 (
    echo ERROR: COM registration failed (see output above).
    pause
    exit /b 1
)
echo [OK] COM add-in registered.

:: ── Write Outlook discovery keys ──────────────────────────────────────────
:: (also done by remitsaver_addin.py, repeated here for clarity)
echo [INFO] Writing Outlook registry key…
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\RemitSaver.Connect" /v "Description"    /t REG_SZ    /d "RemitSaver - Remittance extraction add-in" /f >nul
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\RemitSaver.Connect" /v "FriendlyName"  /t REG_SZ    /d "RemitSaver"                               /f >nul
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\RemitSaver.Connect" /v "LoadBehavior"  /t REG_DWORD /d 3                                           /f >nul
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\RemitSaver.Connect" /v "CommandLineSafe" /t REG_DWORD /d 0                                         /f >nul
echo [OK] Outlook registry key written.

:: ── Create %APPDATA%\RemitSaver directory ─────────────────────────────────
if not exist "%APPDATA%\RemitSaver\" (
    mkdir "%APPDATA%\RemitSaver"
    echo [OK] Created %APPDATA%\RemitSaver\
)

echo.
echo ============================================================
echo  Installation complete!
echo  Restart Outlook – the "RemitSaver" tab will appear in the
echo  ribbon on the main Explorer window.
echo ============================================================
echo.
pause
