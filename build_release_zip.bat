@echo off
setlocal
cd /d "%~dp0"

python "utils\release_zip.py"
if errorlevel 1 (
  echo Ошибка при сборке ZIP.
  exit /b 1
)

echo.
echo Готово. Архив создан в папке "_release".
pause
