@echo off
chcp 65001 >nul
rem 打包給同事用：產出 dist\週報小幫手\，再壓成 dist\週報小幫手.zip
cd /d "%~dp0"
if exist dist rmdir /s /q dist

python clean_template.py || goto :fail

python -m PyInstaller --noconfirm --clean --noconsole --onedir --name 週報小幫手 ^
  --icon Asset\app.ico ^
  --add-data "_dist_template\template.docx;." ^
  --add-data "Asset\app.ico;Asset" ^
  週報小幫手.pyw || goto :fail

copy /y README.txt "dist\週報小幫手\使用說明.txt" >nul
powershell -NoProfile -Command "Compress-Archive -Force -Path 'dist\週報小幫手' -DestinationPath 'dist\週報小幫手.zip'" || goto :fail

echo.
echo 完成：dist\週報小幫手.zip
if /i not "%1"=="nopause" pause
exit /b 0

:fail
echo.
echo 打包失敗，請看上面的錯誤訊息。
if /i not "%1"=="nopause" pause
exit /b 1
