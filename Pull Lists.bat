@echo off
:: Get the directory of the batch file itself
set script_dir=%~dp0

:: Change to the Source directory (one level down)
cd /d "%script_dir%Source"

:: Run the R script with Rscript.exe
"C:\Program Files\R\R-4.3.2\bin\Rscript.exe" "Pull_Lists.R"
