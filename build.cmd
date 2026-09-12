@echo off
rem ============================================================================
rem  QuickPhrase - rebuild the add-in.
rem
rem  Double-click this after changing anything under src\. It will:
rem
rem    1. rebuild QuickPhrase.dotm with Word, exporting the compiled VBA project
rem       to package\word\vbaProject.bin
rem    2. put the Trust Center setting back the way it was
rem
rem  It does not touch git. Nothing is committed and nothing is pushed - review
rem  the result, then commit when you are ready:
rem
rem      git add package/word/vbaProject.bin
rem      git commit -m "build: rebuild compiled VBA project"
rem      git push
rem
rem  Requires: Windows, Microsoft Word. Close Word before running.
rem ============================================================================

setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo ===============================================================
echo  QuickPhrase - build
echo ===============================================================
echo.

rem --- Preflight -------------------------------------------------------------

if not exist "build\build.ps1" (
    echo [ERROR] build\build.ps1 not found. Run this from the repository folder.
    goto :fail
)

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
    echo [ERROR] Word is running. Close every Word window and run this again.
    echo         Word locks the template and can hold stale Trust Center settings.
    goto :fail
)

rem --- Build -----------------------------------------------------------------
rem -EnableVbomTrust sets the registry value that lets the build automate the
rem VBA editor. The restore step below always runs, including when the build
rem fails, so the setting is never left switched on.

echo [1/2] Building with Word...
powershell -NoProfile -ExecutionPolicy Bypass -File "build\build.ps1" -ExportVba -EnableVbomTrust
set "BUILD_RESULT=%ERRORLEVEL%"
echo.

echo [2/2] Restoring the Trust Center setting...
powershell -NoProfile -ExecutionPolicy Bypass -File "build\build.ps1" -DisableVbomTrust
echo.

if not "%BUILD_RESULT%"=="0" (
    echo [ERROR] The build failed.
    goto :fail
)

if not exist "package\word\vbaProject.bin" (
    echo [ERROR] package\word\vbaProject.bin was not produced.
    goto :fail
)

rem --- Report ----------------------------------------------------------------

echo ===============================================================
echo  Build complete.
echo ===============================================================
echo.
echo Template:   dist\QuickPhrase.dotm
echo VBA project: package\word\vbaProject.bin
echo.
echo Install it by double-clicking install\Install.bat
echo.

git rev-parse --is-inside-work-tree >nul 2>&1
if not errorlevel 1 (
    git diff --quiet -- "package/word/vbaProject.bin"
    if errorlevel 1 (
        echo The compiled VBA project changed. To publish it:
        echo.
        echo     git add package/word/vbaProject.bin
        echo     git commit -m "build: rebuild compiled VBA project"
        echo     git push
    ) else (
        echo The compiled VBA project is unchanged - nothing to commit.
        echo Edits to ribbon XML, docs or workflows need no rebuild.
    )
    echo.
)

pause
endlocal
exit /b 0

:fail
echo.
pause
endlocal
exit /b 1
