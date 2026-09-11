@echo off
rem ============================================================================
rem  QuickPhrase - rebuild the VBA project and push it.
rem
rem  Double-click this after changing anything under src\. It will:
rem
rem    1. pull the latest sources
rem    2. rebuild QuickPhrase.dotm with Word, exporting the compiled VBA project
rem    3. put the Trust Center setting back the way it was
rem    4. commit package\word\vbaProject.bin and push
rem
rem  Only the compiled blob is committed. Any other files you changed are
rem  listed and left alone, so this can never push work you were not ready to
rem  share.
rem
rem  Requires: Windows, Microsoft Word, git. Close Word before running.
rem ============================================================================

setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo ===============================================================
echo  QuickPhrase - build and push
echo ===============================================================
echo.

rem --- Preflight -------------------------------------------------------------

where git >nul 2>&1
if errorlevel 1 (
    echo [ERROR] git is not on PATH.
    goto :fail
)

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo [ERROR] This folder is not a git repository:
    echo         %CD%
    goto :fail
)

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
    echo [ERROR] Word is running. Close every Word window and run this again.
    echo         Word locks the template and can hold stale Trust Center settings.
    goto :fail
)

for /f "delims=" %%B in ('git rev-parse --abbrev-ref HEAD') do set "BRANCH=%%B"
echo Branch: !BRANCH!
echo.

rem --- 1. Pull ---------------------------------------------------------------

echo [1/4] Pulling latest changes...
git pull --rebase
if errorlevel 1 (
    echo.
    echo [ERROR] git pull failed. Resolve it by hand, then run this again.
    goto :fail
)
echo.

rem --- 2. Build --------------------------------------------------------------
rem -EnableVbomTrust sets the registry value that lets the build automate the
rem VBA editor. Step 3 always puts it back, including when the build fails.

echo [2/4] Building with Word...
powershell -NoProfile -ExecutionPolicy Bypass -File "build\build.ps1" -ExportVba -EnableVbomTrust
set "BUILD_RESULT=%ERRORLEVEL%"
echo.

echo [3/4] Restoring the Trust Center setting...
powershell -NoProfile -ExecutionPolicy Bypass -File "build\build.ps1" -DisableVbomTrust
echo.

if not "%BUILD_RESULT%"=="0" (
    echo [ERROR] The build failed. Nothing was committed.
    goto :fail
)

if not exist "package\word\vbaProject.bin" (
    echo [ERROR] package\word\vbaProject.bin was not produced.
    goto :fail
)

rem --- 3. Commit and push ----------------------------------------------------

echo [4/4] Committing and pushing...
echo.

git add "package/word/vbaProject.bin"

rem --diff-filter excludes the blob itself, which is already staged.
set "OTHERS="
for /f "delims=" %%F in ('git diff --name-only') do set "OTHERS=1"
if defined OTHERS (
    echo Note: these files are also modified and were NOT committed:
    git diff --name-only
    echo.
)

git diff --cached --quiet
if not errorlevel 1 (
    echo The compiled VBA project is unchanged - nothing to commit.
    echo This is normal if you only edited the ribbon XML, docs or workflows:
    echo those need no rebuild, just a plain git push.
    echo.
    goto :done
)

git commit -m "build: rebuild compiled VBA project"
if errorlevel 1 (
    echo [ERROR] git commit failed.
    goto :fail
)

git push origin "!BRANCH!"
if errorlevel 1 (
    echo.
    echo [ERROR] git push failed.
    goto :fail
)

echo.
echo ===============================================================
echo  Pushed to !BRANCH!.
echo ===============================================================

:done
echo.
echo Built template: dist\QuickPhrase.dotm
echo Install it by double-clicking install\Install.bat
echo.
pause
endlocal
exit /b 0

:fail
echo.
pause
endlocal
exit /b 1
