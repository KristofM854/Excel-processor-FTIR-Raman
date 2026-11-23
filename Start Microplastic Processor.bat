@echo off
setlocal EnableExtensions
title Microplastic Data Processor

rem ==========================================================================
rem  Microplastic Data Processor - desktop launcher
rem
rem  Put this file on your desktop and double-click it. Every start downloads
rem  the latest version of the app from GitHub, so you never need to update
rem  anything yourself. Keep this window open while you use the app; close it
rem  to stop the app.
rem ==========================================================================

rem ---- 1. Find Rscript.exe -------------------------------------------------
rem  a) on the PATH
set "RSCRIPT="
for /f "delims=" %%I in ('where Rscript.exe 2^>nul') do if not defined RSCRIPT set "RSCRIPT=%%I"
if defined RSCRIPT goto :found_r

rem  b) the newest R-x.y.z folder in the usual install locations
rem     (folders are listed by name in ascending order, so the last one wins)
call :find_newest_r "%ProgramFiles%\R"
if defined RSCRIPT goto :found_r
call :find_newest_r "%LOCALAPPDATA%\Programs\R"
if defined RSCRIPT goto :found_r
call :find_newest_r "%ProgramFiles(x86)%\R"
if defined RSCRIPT goto :found_r

rem  c) the install path the R installer writes to the registry
for %%K in (HKLM HKCU) do (
  for /f "tokens=2,*" %%A in ('reg query "%%K\SOFTWARE\R-core\R" /v InstallPath 2^>nul ^| findstr /i "InstallPath"') do (
    if not defined RSCRIPT if exist "%%B\bin\Rscript.exe" set "RSCRIPT=%%B\bin\Rscript.exe"
  )
)
if defined RSCRIPT goto :found_r

if not defined RSCRIPT (
  echo.
  echo  ======================================================================
  echo   R was not found on this computer.
  echo.
  echo   The Microplastic Data Processor needs R. Please install R from
  echo       https://cran.r-project.org/bin/windows/base/
  echo   ^(or ask IT to install it^), then double-click this file again.
  echo  ======================================================================
  echo.
  pause
  exit /b 1
)

:found_r

echo Using R: "%RSCRIPT%"
echo Starting the Microplastic Data Processor (latest version from GitHub)...
echo Keep this window open while you use the app. Close it to stop the app.
echo.

rem ---- 2. Run the R part of this file (below the :R_CODE marker) ----
set "LAUNCHER_FILE=%~f0"
"%RSCRIPT%" -e "x = readLines(Sys.getenv('LAUNCHER_FILE'), warn = FALSE); eval(parse(text = x[-seq_len(grep('^:R_CODE', x)[1])]))"
set "RC=%ERRORLEVEL%"

if "%RC%"=="0" exit /b 0

echo.
echo  ======================================================================
if "%RC%"=="2" (
  echo   No internet connection ^(GitHub could not be reached^).
  echo.
  echo   The app is downloaded fresh from GitHub every time it starts.
  echo   Please check your network / VPN connection and try again.
) else if "%RC%"=="3" (
  echo   The R package "shiny" could not be installed.
  echo.
  echo   Check your internet connection and try again. If it keeps failing,
  echo   ask IT to install the R package "shiny".
) else (
  echo   The app stopped with an error ^(code %RC%^). See the messages above.
  echo   If this keeps happening, send a screenshot of this window to the
  echo   app maintainer.
)
echo  ======================================================================
echo.
pause
exit /b %RC%

rem ---- helper: newest R-* folder below %1 that contains bin\Rscript.exe ----
:find_newest_r
if not exist "%~1\" exit /b 0
for /f "delims=" %%D in ('dir /b /ad /on "%~1\R-*" 2^>nul') do (
  if exist "%~1\%%D\bin\Rscript.exe" set "RSCRIPT=%~1\%%D\bin\Rscript.exe"
)
exit /b 0

rem Everything below is R code, run by the line above. cmd never gets here.
:R_CODE
repo  <- "excel-processor-ftir-raman"
owner <- "KristofM854"
ref   <- "main"

# CRAN mirror and a writable per-user package library (R's own library under
# C:/Program Files is usually read-only).
options(repos = c(CRAN = "https://cloud.r-project.org"))
user_lib <- path.expand(strsplit(Sys.getenv("R_LIBS_USER"), .Platform$path.sep, fixed = TRUE)[[1]][1])
if (!is.na(user_lib) && nzchar(user_lib)) {
  if (!dir.exists(user_lib)) dir.create(user_lib, recursive = TRUE, showWarnings = FALSE)
  if (dir.exists(user_lib)) .libPaths(c(user_lib, .libPaths()))
}

# Internet check with the same download mechanism runGitHub uses.
online <- tryCatch({
  tf <- tempfile()
  suppressWarnings(utils::download.file(
    sprintf("https://raw.githubusercontent.com/%s/%s/%s/DESCRIPTION", owner, repo, ref),
    tf, quiet = TRUE))
  file.exists(tf) && file.info(tf)$size > 0
}, error = function(e) FALSE)
if (!isTRUE(online)) quit(save = "no", status = 2)

if (!requireNamespace("shiny", quietly = TRUE)) {
  message("Installing the R package 'shiny' (first start only, this can take a few minutes)...")
  try(utils::install.packages("shiny"), silent = FALSE)
  if (!requireNamespace("shiny", quietly = TRUE)) quit(save = "no", status = 3)
}

ok <- tryCatch({
  shiny::runGitHub(repo, owner, ref = ref, launch.browser = TRUE)
  TRUE
}, error = function(e) {
  message("\nERROR: ", conditionMessage(e))
  FALSE
})
quit(save = "no", status = if (ok) 0 else 4)
