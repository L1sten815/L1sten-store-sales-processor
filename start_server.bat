@echo off
setlocal
REM === log fallback: if double-click flashes/closes, full output is saved to start_server.log ===
if not defined _SSP_LOGGED (
  set "_SSP_LOGGED=1"
  call "%~f0" > "%~dp0start_server.log" 2>&1
  echo Run log written to: %~dp0start_server.log
  echo (open it with Notepad to see what happened)
  pause
  exit /b
)
REM ============================================================
REM  可移植启动器：发给别人也能用。
REM  用法：把本文件、server.py、store-sales-processor.html 放在同一文件夹，双击即可。
REM  对方需要：Python 3.10+，并已安装 python_calamine 和 openpyxl
REM           （pip install -r requirements.txt）
REM  ============================================================
REM 以「本 bat 自身所在目录」为工作目录（关键：不写死路径）
cd /d "%~dp0"

echo ============================================
echo   Store Sales Processor - starting service
echo ============================================

REM --- 1) 定位 Python 解释器（按优先级回退） ---
set "PYEXE="
REM (a) 随包自带 venv：把含 python_calamine / openpyxl 的 venv 文件夹命名为 venv 放本目录
if exist "%~dp0venv\Scripts\python.exe" set "PYEXE=%~dp0venv\Scripts\python.exe"
REM (b) 原作者机器上的 venv（保持兼容，他人机器此路径不存在会自动跳过）
if not defined PYEXE if exist "C:/Users/71721/.workbuddy/binaries/python/envs/default/Scripts/python.exe" set "PYEXE=C:/Users/71721/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
REM (c) 系统 PATH 中的 python
if not defined PYEXE (
  where python >nul 2>&1
  if not errorlevel 1 set "PYEXE=python"
)
if not defined PYEXE (
  echo [错误] 未找到 Python。请：
  echo   1) 安装 Python 3.10+（勾选 "Add to PATH"）
  echo   2) 在本文件夹打开命令行运行： pip install -r requirements.txt
  echo   3) 再次双击本文件
  pause
  exit /b 1
)
echo 使用 Python: %PYEXE%

REM --- 1.5) 依赖自检 ---
"%PYEXE%" -c "import python_calamine, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo [错误] 缺少依赖 python_calamine / openpyxl。
  echo   请在本文件夹打开命令行运行： pip install -r requirements.txt
  echo   然后再次双击本文件。
  pause
  exit /b 1
)

REM --- 2) 强制结束任何残留的旧服务进程（占着 8000 端口） ---
echo [1/3] 清理旧的 8000 端口占用...
set "KILLED="
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":8000 " ^| findstr LISTENING 2^>nul') do (
    echo     killing PID %%p ...
    taskkill /F /PID %%p >nul 2>&1
    set "KILLED=1"
)
if not defined KILLED echo     no stale process on :8000.
timeout /t 1 >nul

REM --- 3) 启动新的服务 ---
echo [2/3] 启动新的服务进程...
start "Store Sales Processor" "%PYEXE%" server.py

REM --- 等待服务就绪 ---
echo [3/3] 等待服务就绪（约 3 秒）...
timeout /t 3 >nul

echo.
echo ============================================
echo   打开浏览器 http://localhost:8000 使用
echo   关闭 "Store Sales Processor" 窗口可停止服务
echo ============================================
start "" http://localhost:8000

echo.
echo ============================================
echo   服务已在后台启动（独立窗口 "Store Sales Processor"）。
echo   按任意键关闭本窗口即可；不会影响后台服务运行。
echo   若浏览器未自动打开，请手动访问 http://localhost:8000
echo ============================================
pause

endlocal
