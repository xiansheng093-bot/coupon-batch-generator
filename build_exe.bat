@echo off
chcp 65001 >nul 2>&1
title 优惠券批量生成器 - 构建 Windows EXE

echo ============================================
echo   优惠券批量生成器 - Windows EXE 构建脚本
echo ============================================
echo.

:: 检查 Python
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.9+
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

python --version
echo.

:: 检查是否在正确目录
if not exist "coupon_web.py" (
    echo [错误] 当前目录没有 coupon_web.py
    echo 请将此脚本放在项目根目录（与 coupon_web.py 同级）再运行
    pause
    exit /b 1
)

:: 安装依赖
echo [1/3] 安装 Python 依赖...
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)
echo 依赖安装完成！
echo.

:: PyInstaller 打包
echo [2/3] PyInstaller 打包中...
pyinstaller --clean coupon_web.spec
if %ERRORLEVEL% neq 0 (
    echo [错误] 打包失败
    pause
    exit /b 1
)
echo 打包完成！
echo.

:: 生成压缩包
echo [3/3] 生成交付压缩包...
set OUTPUT_DIR=dist\优惠券批量生成器
if exist "%OUTPUT_DIR%" (
    :: 复制示例文件到 dist 目录
    if exist "券码示例.xlsx" copy "券码示例.xlsx" "%OUTPUT_DIR%\"
    if exist "模板示例.png" copy "模板示例.png" "%OUTPUT_DIR%\"

    :: 生成 README
    (
    echo 优惠券批量生成器 / Coupon Batch Generator
    echo ==========================================
    echo.
    echo 使用方法:
    echo   1. 双击 "优惠券批量生成器.exe" 启动服务
    echo   2. 浏览器会自动打开 http://localhost:8888
    echo   3. 按 Ctrl+C 或关闭命令行窗口停止服务
    echo.
    echo 操作流程:
    echo   1. 上传 Excel 券码文件（支持 .xlsx/.xls/.csv）
    echo   2. 选择券码所在列（A-E）
    echo   3. 上传模板图片（支持 PNG/JPG/BMP）
    echo   4. 在预览中拖拽调整二维码和文字位置
    echo   5. 点击 "生成 PDF" 下载多页 PDF
    echo.
    echo 注意: 如果浏览器未自动打开，请手动访问 http://localhost:8888
    ) > "%OUTPUT_DIR%\使用说明.txt"

    :: PowerShell 压缩成 zip
    powershell -Command "Compress-Archive -Path '%OUTPUT_DIR%' -DestinationPath 'dist\优惠券批量生成器-Windows.zip' -Force"
    if %ERRORLEVEL% neq 0 (
        echo [警告] PowerShell 压缩失败，请手动压缩 dist\优惠券批量生成器 目录
    ) else (
        echo.
        echo ============================================
        echo   构建成功！交付文件位于:
        echo   dist\优惠券批量生成器-Windows.zip
        echo ============================================
    )
) else (
    echo [错误] 未找到打包输出目录
    pause
    exit /b 1
)

echo.
echo 完成！可直接将 dist\优惠券批量生成器-Windows.zip 发给客户安装。
pause
