@echo off
chcp 65001 >nul
echo 正在启动店铺任务管理系统...
echo 访问地址: http://localhost:8002
echo 按 Ctrl+C 停止服务
echo.
python app.py
pause
