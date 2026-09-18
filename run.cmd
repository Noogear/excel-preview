@echo off
rem ============================================================================
rem  excel-preview one-click launcher (double-click this file)
rem
rem  Starts the tool only: checks env, installs deps on first run, serves the app,
rem  opens the browser. Closing this window (or Ctrl+C) stops the server and
rem  releases the port - a detached watchdog cleans up even on a forced close.
rem
rem  Options:  -Port 5300   use another port (default 5273, auto-shifts if busy)
rem            -NoBrowser   do not open the browser automatically
rem            -SkipInstall skip the dependency check
rem
rem  Tests / benchmarks are NOT part of this launcher - use npm scripts:
rem    npm test        typecheck + unit tests
rem    npm run e2e     end-to-end tests
rem    npm run bench:parse   performance benchmark
rem
rem  NOTE: keep this file ASCII-only - cmd.exe reads .cmd in the OEM codepage,
rem        so non-ASCII bytes here would be misparsed. Chinese text lives in
rem        tools\start.ps1 (UTF-8 with BOM) instead.
rem ============================================================================
chcp 65001 >nul
setlocal

cd /d "%~dp0"

rem Prefer PowerShell 7 (pwsh); fall back to Windows PowerShell
set "PSEXE=powershell"
where pwsh >nul 2>nul && set "PSEXE=pwsh"

"%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\start.ps1" %*
set "CODE=%ERRORLEVEL%"

if not "%CODE%"=="0" (
  echo.
  echo [launcher] exit code %CODE%
  echo Press any key to close this window...
  pause >nul
)
exit /b %CODE%
