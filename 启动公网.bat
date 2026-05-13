@echo off
echo ============================================
echo   Gamma PPTX 处理工具 + 公网链接
echo ============================================

echo.
echo [1/2] 启动本地服务...
start /b "pptx-server" "D:\msy64\ucrt64\bin\python.exe" "C:\Users\Lmy-pc\Desktop\gamma-tool\pptx_tool.py"
timeout /t 2 /nobreak >nul

echo [2/2] 创建公网隧道...
echo.
echo 公网链接会在下方显示，复制那个 https://xxx.lhr.life 的链接即可
echo 按 Ctrl+C 可以停止
echo ============================================
echo.

ssh -o StrictHostKeyChecking=no -R 80:localhost:8999 nokey@localhost.run
