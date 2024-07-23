@echo off
setlocal

:: 设置打包输出目录
set DIST_DIR=dist

:: 使用 PyInstaller 打包应用程序
pyinstaller -F --add-data "templates;templates" .\app.py

:: 检查打包是否成功
IF %ERRORLEVEL% NEQ 0 (
    echo PyInstaller打包失败
    exit /b %ERRORLEVEL%
)

:: 复制 templates 目录到打包后的 dist 目录
xcopy /E /I /Y "templates" "%DIST_DIR%\templates"

:: 检查复制是否成功
IF %ERRORLEVEL% NEQ 0 (
    echo 复制 templates 目录失败
    exit /b %ERRORLEVEL%
)

echo 打包和复制完成
endlocal