@echo off
REM DOCX to PDF/A Converter Batch Script
REM Usage: convert.bat "path\to\document.docx"

setlocal enabledelayedexpansion

REM Check if input file is provided
if "%~1"=="" (
    echo.
    echo Usage: convert.bat "path\to\document.docx"
    echo.
    echo Example:
    echo   convert.bat "C:\Documents\myfile.docx"
    echo.
    exit /b 1
)

REM Get the input DOCX path
set "DOCX_INPUT=%~1"

REM Check if file exists
if not exist "!DOCX_INPUT!" (
    echo Error: File not found - !DOCX_INPUT!
    exit /b 1
)

REM Check if it's a DOCX file
if /i not "!DOCX_INPUT:~-5!"==".docx" (
    echo Error: File must be a .docx file
    exit /b 1
)

REM Get the directory and filename without extension
set "DOCX_DIR=%~dp1"
set "FILENAME=%~n1"

REM Create output PDF path (same directory, same filename, .pdf extension)
set "PDF_OUTPUT=!DOCX_DIR!!FILENAME!.pdf"

REM Check if virtual environment exists
if not exist ".venv\Scripts\python.exe" (
    echo Error: Virtual environment not found. Please run:
    echo   python -m venv .venv
    echo   .venv\Scripts\pip install -r requirements.txt
    exit /b 1
)

REM Find Ghostscript path
set "GS_PATH="
for /D %%G in ("C:\Program Files\gs\gs*") do (
    if exist "%%G\bin\gswin64c.exe" (
        set "GS_PATH=%%G\bin\gswin64c.exe"
    )
)

REM Run the converter
echo.
echo Converting: !DOCX_INPUT!
echo Output:     !PDF_OUTPUT!
echo.

if "!GS_PATH!"=="" (
    .venv\Scripts\python.exe src\convert.py "!DOCX_INPUT!" "!PDF_OUTPUT!" --bookmarks headings
) else (
    .venv\Scripts\python.exe src\convert.py "!DOCX_INPUT!" "!PDF_OUTPUT!" --bookmarks headings --gs-path "!GS_PATH!"
)

if !errorlevel! neq 0 (
    echo.
    echo Error: Conversion failed with exit code !errorlevel!
    exit /b !errorlevel!
)

echo.
echo Success! PDF/A created: !PDF_OUTPUT!
echo.
