@echo off
chcp 936 >nul 2>&1

echo ================================================
echo   文档搜索索引 v1.0 - 一键打包工具
echo ================================================
echo.

REM --- Step 1: Find Python ---
set PYTHON_CMD=

python --version >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=python
if not errorlevel 1 goto found_python

py --version >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=py
if not errorlevel 1 goto found_python

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" goto found_python
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" goto found_python
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python310\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" goto found_python
if exist "%LOCALAPPDATA%\Programs\Python\Python39\python.exe" set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python39\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python39\python.exe" goto found_python
if exist "C:\Python312\python.exe" set PYTHON_CMD=C:\Python312\python.exe
if exist "C:\Python312\python.exe" goto found_python
if exist "C:\Python311\python.exe" set PYTHON_CMD=C:\Python311\python.exe
if exist "C:\Python311\python.exe" goto found_python
if exist "C:\Python310\python.exe" set PYTHON_CMD=C:\Python310\python.exe
if exist "C:\Python310\python.exe" goto found_python

echo [错误] 未找到 Python！
echo.
echo 请先安装 Python 3.8 或更高版本：
echo   https://www.python.org/downloads/
echo.
echo 安装时请勾选 Add Python to PATH 选项！
echo.
pause
exit /b 1

:found_python
echo [1/4] 找到 Python：
%PYTHON_CMD% --version
echo.

REM --- Step 2: Install dependencies ---
echo [2/4] 正在安装依赖库，首次需要几分钟，请耐心等待...
%PYTHON_CMD% -m pip install --upgrade pip -q --no-warn-script-location
%PYTHON_CMD% -m pip install python-docx openpyxl python-pptx pdfplumber xlrd jieba PyQt5 pyinstaller -q --no-warn-script-location
if errorlevel 1 goto install_error
echo   依赖安装完成！
echo.

REM --- Step 3: Clean old build ---
if exist "dist\DocFinder_v1.0.exe" del /f /q "dist\DocFinder_v1.0.exe"
if exist "build" rmdir /s /q "build"

REM --- Step 4: Package with PyInstaller ---
echo [3/4] 开始打包，需要 3-8 分钟，请勿关闭窗口...
echo.

%PYTHON_CMD% -m PyInstaller ^
    --clean ^
    --onefile ^
    --windowed ^
    --name DocFinder_v1.0 ^
    --hidden-import jieba ^
    --hidden-import jieba.analyse ^
    --hidden-import jieba.posseg ^
    --hidden-import docx ^
    --hidden-import docx.oxml ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.styles ^
    --hidden-import pptx ^
    --hidden-import pptx.util ^
    --hidden-import pdfplumber ^
    --hidden-import xlrd ^
    --hidden-import PyQt5 ^
    --hidden-import PyQt5.QtWidgets ^
    --hidden-import PyQt5.QtCore ^
    --hidden-import PyQt5.QtGui ^
    --collect-data jieba ^
    --exclude-module matplotlib ^
    --exclude-module numpy ^
    --exclude-module scipy ^
    --exclude-module pandas ^
    --exclude-module tkinter ^
    --noupx ^
    main.py

if errorlevel 1 goto pack_error

REM --- Done ---
echo.
echo [4/4] 打包完成！
echo.
echo ================================================
echo   成功！输出文件：dist\DocFinder_v1.0.exe
echo.
echo   使用方法：
echo   1. 将 exe 复制到任意位置，双击运行
echo   2. 菜单 索引-新建索引 开始建立索引
echo   3. 在搜索框输入关键词即可搜索
echo ================================================
echo.
explorer dist
pause
exit /b 0

:install_error
echo.
echo [错误] 依赖安装失败！请检查网络连接后重试。
echo.
pause
exit /b 1

:pack_error
echo.
echo [错误] 打包失败！
echo 常见原因：
echo   1. 杀毒软件拦截，请暂时关闭后重试
echo   2. 磁盘空间不足
echo   3. 依赖库安装不完整
echo.
pause
exit /b 1
