@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================
echo   投保数据汇总工具 - Windows 打包脚本
echo ==========================================
echo.

echo [1/3] 安装依赖（首次需要联网，约1-3分钟）...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败，请确认已安装 Python 并加入环境变量
    pause
    exit /b 1
)

echo.
echo [2/3] 清理旧的打包产物...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [3/3] 正在打包，需要 2-5 分钟，请勿关闭窗口...
python -m PyInstaller --noconfirm --onedir --windowed --clean ^
  --name "投保数据汇总工具" ^
  --add-data "templates;templates" ^
  --collect-all webview ^
  --exclude-module matplotlib ^
  --exclude-module scipy ^
  --exclude-module PyQt5 ^
  --exclude-module IPython ^
  --exclude-module notebook ^
  --exclude-module pytest ^
  insurance_web.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请把上面的报错信息发给维护人员
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   打包完成！
echo.
echo   程序位置: dist\投保数据汇总工具\
echo   请右键把该文件夹压缩成 ZIP，发给同事
echo   同事解压后双击里面的 exe 即可使用
echo ==========================================
echo.
pause
