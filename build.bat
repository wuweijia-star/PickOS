@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title PickOS 打包工具（Python 3.12 · onedir）
cd /d "%~dp0"

echo ============================================
echo    PickOS 班级管理系统 - 打包工具
echo    打包环境：Python 3.12  ·  onedir 模式
echo ============================================
echo.

:: ============ 配置区 ============
set "APP_NAME=PickOS"
set "MAIN_FILE=main.pyw"
set "ICON_FILE=icon.ico"
set "PYTHON_CMD=py -3.12"
:: ★ 已删除硬编码 PYTHON312_DLL，改为下方动态获取
:: ================================

:: ---- 1. 检查 Python 3.12 ----
echo [1/6] 检查 Python 3.12...
%PYTHON_CMD% --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python 3.12
    pause
    exit /b 1
)
%PYTHON_CMD% --version
echo.

:: ★ 动态获取 python312.dll 的真实路径（仅用于启动前检查与信息输出）
::   PyInstaller 6.x 的 onedir 模式会自动把 python312.dll 收进 _internal/，
::   这里不再需要 --add-binary。
for /f "delims=" %%i in ('%PYTHON_CMD% -c "import sys,os;print(os.path.dirname(sys.executable))"') do set "PYTHON_DIR=%%i"
set "PYTHON312_DLL=%PYTHON_DIR%\python312.dll"
if not exist "%PYTHON312_DLL%" (
    echo [错误] 找不到 python312.dll：%PYTHON312_DLL%
    echo 请确认 Python 3.12 安装完整
    pause
    exit /b 1
)
echo Python 目录：%PYTHON_DIR%
echo python312.dll：%PYTHON312_DLL%
echo.

:: ---- 2. 检查主文件 ----
if not exist "%MAIN_FILE%" (
    for %%f in (*.pyw) do (
        set "MAIN_FILE=%%f"
        goto :got_main
    )
    echo [错误] 未找到 .pyw 文件
    pause
    exit /b 1
)
:got_main
echo 主文件：%MAIN_FILE%
echo.

:: ---- 3. 安装依赖 ----
echo [2/6] 检查依赖...

:: ★ 基础依赖：pypi 失败自动回退清华镜像
%PYTHON_CMD% -m pip install PyQt6 openpyxl pandas psutil lunardate -i https://pypi.org/simple/
if errorlevel 1 (
    echo [提示] pypi 失败，改用清华镜像重试...
    %PYTHON_CMD% -m pip install PyQt6 openpyxl pandas psutil lunardate -i https://pypi.tuna.tsinghua.edu.cn/simple
)
if errorlevel 1 (
    echo [错误] 基础依赖安装失败
    pause
    exit /b 1
)

:: ★ qfluentwidgets：同样加镜像回退 + 错误检查
%PYTHON_CMD% -m pip install PyQt6-Fluent-Widgets -i https://pypi.org/simple/
if errorlevel 1 (
    echo [提示] pypi 失败，改用清华镜像重试...
    %PYTHON_CMD% -m pip install PyQt6-Fluent-Widgets -i https://pypi.tuna.tsinghua.edu.cn/simple
)
if errorlevel 1 (
    echo [错误] PyQt6-Fluent-Widgets 安装失败
    pause
    exit /b 1
)

:: ★ pyinstaller：同样加镜像回退 + 错误检查
%PYTHON_CMD% -m pip install pyinstaller -i https://pypi.org/simple/
if errorlevel 1 (
    echo [提示] pypi 失败，改用清华镜像重试...
    %PYTHON_CMD% -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
)
if errorlevel 1 (
    echo [错误] PyInstaller 安装失败
    pause
    exit /b 1
)

echo 验证依赖...
%PYTHON_CMD% -c "import PyQt6; import qfluentwidgets; import pandas; import openpyxl; import psutil; import lunardate; print('[OK] 依赖正常')"
if errorlevel 1 (
    echo [错误] 依赖导入失败
    pause
    exit /b 1
)
echo.

:: ---- 4. 清理 ----
echo [3/6] 清理旧文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "%APP_NAME%.spec" del /q "%APP_NAME%.spec"
echo.

:: ---- 5. 打包（onedir）----
echo [4/6] 开始打包（约 3-8 分钟）...
echo.

set "ICON_ARG="
if exist "%ICON_FILE%" (
    set ICON_ARG=--icon="%ICON_FILE%"
    echo 使用图标：%ICON_FILE%
) else (
    echo [提示] 无图标，使用默认
)
echo.

:: ★ 说明：
::   · 去掉了 --add-binary "%PYTHON312_DLL%;."（PyInstaller hook 自动处理）
::   · 去掉了 --collect-all PyQt6（会拖进 QtOpenGL/QtMultimedia 等，体积暴增）
::   · 去掉了 --noupx（PyInstaller 6.x 默认就不带 UPX）
::   · 调试阶段可临时把 --windowed 改成 --console，方便看报错
::
:: ★ 可选：如果想让"新手引导 → 创建桌面快捷方式"在打包后也能用，
::   请先执行：  py -3.12 -m pip install pywin32
::   然后在本文件末尾追加以下 hidden-import：
::       --hidden-import=win32com ^
::       --hidden-import=win32com.client ^
::       --hidden-import=pythoncom ^
::   （不加也没关系，程序会走 try/except 分支，只是该功能会提示"缺少 pywin32"）

%PYTHON_CMD% -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --windowed ^
    --onedir ^
    --name "%APP_NAME%" ^
    !ICON_ARG! ^
    --collect-all qfluentwidgets ^
    --collect-data pandas ^
    --collect-data numpy ^
    --collect-submodules openpyxl ^
    --collect-submodules psutil ^
    --hidden-import=lunardate ^
    --hidden-import=darkdetect ^
    --hidden-import=ctypes ^
    --hidden-import=ctypes.wintypes ^
    --hidden-import=sqlite3 ^
    --hidden-import=PyQt6.QtSvg ^
    --hidden-import=PyQt6.QtSvgWidgets ^
    --hidden-import=PyQt6.QtNetwork ^
    --hidden-import=PyQt6.QtPrintSupport ^
    "%MAIN_FILE%"

if errorlevel 1 (
    echo.
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo [√] 打包成功
echo.
echo ============================================
echo   输出目录：dist\%APP_NAME%\
echo   可执行文件：%APP_NAME%.exe
echo ============================================
echo.
echo [重要提示]
echo   1. onedir 模式：双击即开，启动很快，无临时目录警告
echo   2. 分发时把整个 dist\%APP_NAME%\ 文件夹一起拷走
echo   3. 数据仍会在 exe 同目录创建 user\ 文件夹
echo   4. 若首次启动闪退，请把 --windowed 换成 --console 重打一次看报错
echo   5. 不要把 exe 放到 C:\Program Files 这类只读目录，否则 user\ 无法写入
echo.

choice /c YN /m "是否立即打开输出目录"
if errorlevel 2 goto :end
if errorlevel 1 explorer "dist\%APP_NAME%"

:end
pause
endlocal
