@echo off
chcp 65001 >nul
REM ===== Lưu toàn bộ thay đổi lên GitHub trong 1 cái bấm =====
REM Bấm đúp vào file này là: thêm hết -> commit -> push.
cd /d "%~dp0"

echo.
echo ==== Cac thay doi hien tai ====
git status -s

echo.
set /p msg="Nhap mo ta thay doi (Enter de dung mac dinh): "
if "%msg%"=="" set msg=Cap nhat noi dung

git add -A
git commit -m "%msg%"
git push

echo.
echo ==== XONG. Vercel va Render se tu deploy lai. ====
pause
