@if (@a==@b) @end /*
@echo off
:: Batch portion
REM Parse input arguments
FOR /f "tokens=1,2,3,4,5 delims=:" %%a IN (%1) DO (
    set Key=%%c
    set Val=%%d
    set Art=%%b
    set name=%%e
)
 
REM Run the JScript portion to select the folder
for /f "delims=" %%I in ('cscript /nologo /e:jscript "%~f0"') do (
    set thefiles=%%I
)
 
REM Split the report into sections based on placeholders
REM Each placeholder starts a new section with a specified orientation
setlocal enabledelayedexpansion
set counter=1

REM Temporary folder for individual PDFs
set tempdir=%~dp0temp_pdf
if not exist "%tempdir%" mkdir "%tempdir%"
set URL=https://gi.grassone.be/?p=%Art%
 
REM Read the input HTML or report content
REM Using a PowerShell command here for better parsing
echo %URL%


REM  [Parameter(Mandatory = $true, HelpMessage = 'authcookies_key')]
REM  [string]$authcookies_key,
REM  [Parameter(Mandatory = $true, HelpMessage = 'authcookies_value')]
REM  [string]$authcookies_value,
REM  [Parameter(Mandatory = $true, HelpMessage = 'root URL')]
REM  [string]$root_URL,
REM  [Parameter(Mandatory = $true, HelpMessage = 'post to print')]
REM  [string]$article,
REM  [Parameter(Mandatory = $true, HelpMessage = 'outputPDF')]
REM  [string]$outputPDF,
REM  [Parameter(Mandatory = $true, HelpMessage = 'Page counter postscript : C:\temp_pdf\pagecount.ps')]
REM  [string]$pageCounterPS,
REM  [Parameter(Mandatory = $true, HelpMessage = 'Log folder : C:\temp_pdf\ ')]
REM  [string]$Log_folder 
  
powershell  -ExecutionPolicy Bypass -File "%~dp0htmltoPDF.ps1" -authcookies_key "%Key%" -authcookies_value "%Val%" -root_URL "https://gi.grassone.be" -article "%Art%" -outputPDF "%thefiles%\%name%.pdf" -pageCounterPS "%~dp0pagecount.ps" -Log_folder "%tempdir%"
 
 
 
REM Open the resulting PDF
"%thefiles%\%name%.pdf"

goto :EOF

:: JScript portion */

var shl = new ActiveXObject("Shell.Application");
var folder = shl.BrowseForFolder(0, "Please choose a folder.", 0, 0x00);
WSH.Echo(folder ? folder.self.path : '')