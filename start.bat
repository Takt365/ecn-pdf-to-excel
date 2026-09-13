@echo off
chcp 65001 >nul 2>&1
title ECN PDF to Excel

echo ========================================
echo   ECN PDF to Excel Converter
echo ========================================
echo.
echo  1. Convert single PDF
echo  2. Batch convert folder
echo  3. Exit
echo.
set /p CHOICE="Select (1/2/3): "

if "%CHOICE%"=="1" goto SINGLE
if "%CHOICE%"=="2" goto BATCH
if "%CHOICE%"=="3" exit /b

:SINGLE
set /p PDFPATH="PDF file path: "
node ecn_pdf_to_excel.js "%PDFPATH%"
pause
exit /b

:BATCH
set /p FOLDER="Folder path containing PDFs: "
node batch.js "%FOLDER%"
pause
exit /b