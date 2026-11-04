@echo off
setlocal EnableExtensions
cd /d "%~dp0"

rem ================================================================
rem  AutoReport build - produces dist\Report.exe.
rem
rem  Self-contained: no build.py, no .spec to maintain. PyInstaller is
rem  called straight on report.py; the .spec it writes is a throwaway
rem  in build\ and is regenerated on every run (--specpath).
rem
rem  The .docx templates and asset\ cannot be bundled inside a
rem  --onefile exe: report.py resolves them next to the exe
rem  (APP_DIR), not inside the temporary _MEIPASS folder. They are
rem  copied into dist\ after the freeze instead.
rem
rem  Every run writes build_log.txt with the tool versions and the
rem  result. Set AUTOREPORT_NO_PAUSE=1 to skip the final key press.
rem ================================================================

rem %~dp0 always ends with a backslash, so paths below add no separator.
set "ROOT=%~dp0"
set "LOG=%ROOT%build_log.txt"
set "VENV=%ROOT%.venv"
set "VPY=%VENV%\Scripts\python.exe"
set "DIST=%ROOT%dist"
set "WORK=%ROOT%build"
set "APPNAME=Report"

echo ================================================================
echo  AutoReport build
echo ================================================================
echo  folder : %ROOT%
echo.

> "%LOG%" echo ============================================================
>> "%LOG%" echo AutoReport build
>> "%LOG%" echo started : %DATE% %TIME%
>> "%LOG%" echo folder  : %ROOT%
>> "%LOG%" echo ============================================================

rem ---- Locate Python -------------------------------------------------
rem Newest supported interpreter first: the pinned wheels are validated
rem on 3.11/3.12, and a machine that also has 3.15+ would otherwise be
rem picked by a bare "python".
set "PY="
call :find_python "py -3.12"
call :find_python "py -3.11"
call :find_python "py -3"
call :find_python "python"
if not defined PY goto :no_python

echo  python : %PY%
echo.
%PY% -c "import sys; print('  version : ' + sys.version.split()[0]); print('  exe     : ' + sys.executable)"
if errorlevel 1 goto :failed
>> "%LOG%" echo python  : %PY%

rem ---- 1. Virtual environment ---------------------------------------
echo.
echo [1/6] Virtual environment
if exist "%VPY%" (
    echo    reusing %VENV%
) else (
    echo    creating %VENV%
    %PY% -m venv "%VENV%"
    if errorlevel 1 goto :failed
)
if not exist "%VPY%" goto :failed
>> "%LOG%" echo venv    : %VPY%

rem ---- 2. Runtime requirements ---------------------------------------
echo.
echo [2/6] Runtime requirements
"%VPY%" -m pip install --upgrade pip --disable-pip-version-check --quiet
"%VPY%" -m pip install -r "%ROOT%requirements.txt" --only-binary :all: --disable-pip-version-check
if errorlevel 1 goto :failed

rem ---- 3. PyInstaller -------------------------------------------------
echo.
echo [3/6] PyInstaller
"%VPY%" -m pip install -r "%ROOT%requirements-build.txt" --only-binary :all: --disable-pip-version-check
if errorlevel 1 goto :failed
for /f "delims=" %%V in ('"%VPY%" -m PyInstaller --version') do set "PYIVER=%%V"
>> "%LOG%" echo pyinstaller : %PYIVER%

rem ---- 4. Sanity checks ----------------------------------------------
echo.
echo [4/6] Checking report.py
"%VPY%" -m py_compile "%ROOT%report.py"
if errorlevel 1 goto :failed
"%VPY%" -c "import report"
if errorlevel 1 goto :failed
echo    syntax and imports ok

rem ---- 5. Freeze ------------------------------------------------------
echo.
echo [5/6] Freezing to dist\%APPNAME%.exe
if exist "%DIST%" rmdir /s /q "%DIST%"
if exist "%WORK%" rmdir /s /q "%WORK%"
mkdir "%WORK%"
if not exist "%ROOT%asset\app.ico" (
    echo [ERROR] asset\app.ico is missing - that is the exe icon.
    goto :failed
)
copy /y "%ROOT%asset\app.ico" "%WORK%\app.ico" >nul

"%VPY%" -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name %APPNAME% ^
    --icon "%WORK%\app.ico" ^
    --collect-data tzdata ^
    --exclude-module tkinter ^
    --exclude-module numpy ^
    --exclude-module pandas ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module IPython ^
    --distpath "%DIST%" ^
    --workpath "%WORK%" ^
    --specpath "%WORK%" ^
    "%ROOT%report.py"

if errorlevel 1 goto :failed
if not exist "%DIST%\%APPNAME%.exe" (
    echo [ERROR] PyInstaller finished but dist\%APPNAME%.exe is missing.
    goto :failed
)

rem ---- 6. Files the app reads next to the exe -------------------------
echo.
echo [6/6] Copying templates and assets into dist
if exist "%ROOT%Report Template.docx" (
    copy /y "%ROOT%Report Template.docx" "%DIST%\" >nul
    echo    copied Report Template.docx
) else (
    echo [WARN] Report Template.docx not found - Status Report will fail until it is in dist\
)
if exist "%ROOT%POT_Template.docx" (
    copy /y "%ROOT%POT_Template.docx" "%DIST%\" >nul
    echo    copied POT_Template.docx
) else (
    echo [WARN] POT_Template.docx not found - Create POT will fail until it is in dist\
)
if exist "%ROOT%asset" (
    xcopy /e /i /y /q "%ROOT%asset" "%DIST%\asset\" >nul
    echo    copied asset\
)

set "SIZE=0"
for %%A in ("%DIST%\%APPNAME%.exe") do set "SIZE=%%~zA"
set /a "SIZEMB=%SIZE% / 1048576" 2>nul

echo.
echo ================================================================
echo  DONE
echo    exe  : %DIST%\%APPNAME%.exe
echo    size : %SIZEMB% MB
echo    log  : %LOG%
echo ================================================================
>> "%LOG%" echo finished: %DATE% %TIME%
>> "%LOG%" echo exe     : %DIST%\%APPNAME%.exe
>> "%LOG%" echo size    : %SIZE% bytes
>> "%LOG%" echo ============================================================

if not defined AUTOREPORT_NO_PAUSE pause
exit /b 0

:failed
echo.
echo [FAILED] Build failed - read the messages above.
>> "%LOG%" echo FAILED at %TIME%
>> "%LOG%" echo ============================================================
if not defined AUTOREPORT_NO_PAUSE pause
exit /b 1

:no_python
echo [ERROR] No Python found. Install Python 3.12 from python.org and tick "Add python.exe to PATH".
if not defined AUTOREPORT_NO_PAUSE pause
exit /b 1

:find_python
rem %~1 is a full command line such as "py -3.12"; the first one that
rem answers --version wins and is remembered in PY.
if defined PY exit /b 0
%~1 --version >nul 2>&1
if errorlevel 1 exit /b 0
set "PY=%~1"
exit /b 0
