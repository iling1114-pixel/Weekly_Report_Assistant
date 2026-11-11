@echo off
chcp 65001 >nul
rem 發布新版到 GitHub：先把 report.py 的 VERSION 改大，再雙擊這個檔
cd /d "%~dp0"

for /f %%v in ('python -c "import report; print(report.VERSION)"') do set VER=%%v
if "%VER%"=="" goto :fail
echo 準備發布 v%VER%

gh release view v%VER% >nul 2>&1 && (
  echo v%VER% 已經發布過了，請先把 report.py 的 VERSION 改大。
  goto :fail
)

set /p NOTES=這版改了什麼（一行，會顯示在下載頁）：
if "%NOTES%"=="" set NOTES=v%VER%

git add -A || goto :fail
git commit -m "v%VER%: %NOTES%"
git push || goto :fail

call build.bat nopause || goto :fail

rem 檔名用英文（GitHub 會把中文檔名改掉），顯示名稱用中文
copy /y "dist\週報小幫手.zip" "dist\WeeklyReportHelper-v%VER%.zip" >nul || goto :fail
gh release create v%VER% "dist\WeeklyReportHelper-v%VER%.zip#週報小幫手 v%VER%.zip" --title "v%VER%" --notes "%NOTES%" || goto :fail

echo.
echo 完成！同事打開程式就會看到「有新版本 v%VER%」提醒。
pause
exit /b 0

:fail
echo.
echo 發布失敗，請看上面的錯誤訊息。
pause
exit /b 1
