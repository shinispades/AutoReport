@echo off
setlocal EnableExtensions
cd /d "%~dp0"

rem ================================================================
rem  AutoReport build - thin launcher.
rem
rem  The actual work (venv, pip, PyInstaller, logging, error reporting)
rem  lives in build.py: cmd.exe quoting and error handling are exactly
rem  where builds break, so this file only locates Python and hands over.
rem
rem  Every run writes build_log.txt. On failure the last lines of the
rem  failing step are printed here, and the full log is on disk.
rem
rem  Set AUTOREPORT_NO_PAUSE=1 to keep the window from waiting for a key
rem  (for CI, or when calling this from another script).
rem ================================================================

echo ================================================================
echo  AutoReport build
echo ================================================================

rem Prefer a supported interpreter explicitly, newest first. `py -3` picks the
rem newest installed Python, which on a machine that also has 3.15+ would land
rem on an interpreter the pinned wheels do not cover.
set "PY="
if not defined PY py -3.12 --version >nul 2>&1
if not defined PY if not errorlevel 1 set "PY=py -3.12"
if not defined PY py -3.11 --version >nul 2>&1
if not defined PY if not errorlevel 1 set "PY=py -3.11"
if not defined PY py -3 --version >nul 2>&1
if not defined PY if not errorlevel 1 set "PY=py -3"
if not defined PY python --version >nul 2>&1
if not defined PY if not errorlevel 1 set "PY=python"
if not defined PY echo [ERROR] No Python found. Install Python 3.12 from python.org and tick "Add python.exe to PATH".
if not defined PY if not defined AUTOREPORT_NO_PAUSE pause
if not defined PY exit /b 1

%PY% "%~dp0build.py"
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" echo [FAILED] Build failed - see the output above and build_log.txt.
if not "%RC%"=="0" if not defined AUTOREPORT_NO_PAUSE pause
if "%RC%"=="0" echo [OK] Build finished.
if "%RC%"=="0" if not defined AUTOREPORT_NO_PAUSE pause
exit /b %RC%
