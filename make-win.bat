@echo off
setlocal

set SCRIPT_DIR=%~dp0
set WIN_DIR=%SCRIPT_DIR%windows-xlsx2csv

if not exist "%WIN_DIR%" mkdir "%WIN_DIR%"

cargo build --release
if errorlevel 1 exit /b 1

copy /Y "target\release\xlsx2csv.exe" "%WIN_DIR%\"
copy /Y "README.md" "%WIN_DIR%\"
copy /Y "README-*.md" "%WIN_DIR%\"
copy /Y "LICENSE" "%WIN_DIR%\"
copy /Y "Cargo.toml" "%WIN_DIR%\"
copy /Y "xlsx2csv-win.toml" "%WIN_DIR%\xlsx2csv.toml"

cd /d "%WIN_DIR%"
powershell -NoProfile -Command "Compress-Archive -Force -Path * -DestinationPath '%SCRIPT_DIR%windows-xlsx2csv.zip'"

echo 完了: windows-xlsx2csv.zip
endlocal
