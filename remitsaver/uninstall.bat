@echo off
:: ============================================================
:: RemitSaver – COM Add-in Unregistration
:: Run this script as Administrator.
:: ============================================================

setlocal EnableDelayedExpansion

echo ============================================================
echo  RemitSaver Uninstaller
echo ============================================================
echo.

:: ── Verify Administrator ─────────────────────────────────────────────────
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo ERROR: This script must be run as Administrator.
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
echo WARNING: Python not found – will only remove registry keys.
goto :remove_registry
:found_python

:: ── Unregister COM server ─────────────────────────────────────────────────
echo [INFO] Unregistering RemitSaver COM server…
%PYTHON% "%~dp0remitsaver_addin.py" --unregister
if %errorLevel% NEQ 0 (
    echo WARNING: COM unregistration returned an error (may already be removed).
)
echo [OK] COM server unregistered.

:remove_registry
:: ── Remove Outlook discovery key ──────────────────────────────────────────
echo [INFO] Removing Outlook registry key…
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\RemitSaver.Connect" /f >nul 2>&1
echo [OK] Outlook registry key removed.

:: ── Remove COM CLSID key (belt-and-braces) ────────────────────────────────
reg delete "HKCU\Software\Classes\CLSID\{6D8B3E2A-4F1C-4A7D-9B5E-2C8F1A3D6E9B}" /f >nul 2>&1
reg delete "HKCU\Software\Classes\RemitSaver.Connect" /f >nul 2>&1

echo.
echo ============================================================
echo  Uninstallation complete.
echo  Configuration files in %%APPDATA%%\RemitSaver\ were kept.
echo  Delete that folder manually if you want a full clean.
echo ============================================================
echo.
pause
