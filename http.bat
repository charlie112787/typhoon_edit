@echo off

echo.
echo =================================
echo 正在啟動伺服器...
echo =================================
echo.

echo 執行指令：python -m http.server 8000
python -m http.server 8000
echo.
echo 伺服器已停止。
echo.

pause