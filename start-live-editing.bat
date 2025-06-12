@echo off
REM Enhanced Word MCP Server - Live Editing Startup Script (Windows)
REM This script automatically starts both the MCP server and Word Add-in development server

setlocal EnableDelayedExpansion

set "PROJECT_DIR=%~dp0"
set "ADDIN_DIR=%PROJECT_DIR%word-live-addin"

echo.
echo ============================================================
echo 🚀 Enhanced Word MCP Server - Live Editing Setup
echo ============================================================
echo Project directory: %PROJECT_DIR%

REM Function to check if a port is in use
:check_port
netstat -an | findstr ":%1 " | findstr "LISTENING" >nul 2>&1
exit /b %errorlevel%

REM Start MCP Server
echo.
echo 📡 Starting MCP Server (WebSocket on port 8765)...

call :check_port 8765
if %errorlevel% == 0 (
    echo ⚠️  Port 8765 is already in use - MCP server may already be running
    echo    If you need to restart, please stop the existing server first
    pause
    exit /b 1
)

cd /d "%PROJECT_DIR%"

REM Check if virtual environment exists and activate it
if exist "venv\Scripts\activate.bat" (
    echo    Activating virtual environment...
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    echo    Activating virtual environment...
    call .venv\Scripts\activate.bat
)

REM Install dependencies if needed
if exist "requirements.txt" (
    echo    Installing Python dependencies...
    pip install -r requirements.txt >nul 2>&1
)

REM Start MCP server in background
echo    Launching MCP server...
start "MCP Server" /min python -m word_document_server.main

REM Wait for server to start
timeout /t 3 /nobreak >nul

REM Check if MCP server started successfully
call :check_port 8765
if %errorlevel% == 0 (
    echo ✅ MCP Server started successfully
    echo    WebSocket endpoint: ws://localhost:8765
) else (
    echo ❌ Failed to start MCP server
    pause
    exit /b 1
)

REM Start Word Add-in Development Server
echo.
echo 🔧 Starting Word Add-in Development Server (HTTPS on port 3000)...

call :check_port 3000
if %errorlevel% == 0 (
    echo ⚠️  Port 3000 is already in use - Add-in server may already be running
    echo    If you need to restart, please stop the existing server first
    pause
    exit /b 1
)

cd /d "%ADDIN_DIR%"

REM Check if Node.js is installed
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Node.js is not installed. Please install Node.js and npm first.
    echo    Visit: https://nodejs.org/
    pause
    exit /b 1
)

REM Install npm dependencies if node_modules doesn't exist
if not exist "node_modules" (
    echo    Installing npm dependencies...
    call npm install
)

REM Start Add-in development server in background
echo    Launching Add-in development server...
start "Add-in Server" /min npm start

REM Wait for server to start
timeout /t 5 /nobreak >nul

REM Check if Add-in server started successfully
call :check_port 3000
if %errorlevel% == 0 (
    echo ✅ Add-in Development Server started successfully
    echo    HTTPS endpoint: https://localhost:3000
) else (
    echo ❌ Failed to start Add-in development server
    pause
    exit /b 1
)

echo.
echo ============================================================
echo 🎉 Live Editing System Ready!
echo ============================================================
echo.
echo 📋 Next Steps:
echo 1. Open Microsoft Word
echo 2. Go to Insert → My Add-ins → Upload My Add-in
echo 3. Select: %ADDIN_DIR%\manifest.xml
echo 4. Open a document and the Add-in will auto-connect!
echo.
echo 🔗 Endpoints:
echo    • MCP Server: ws://localhost:8765
echo    • Add-in Dev: https://localhost:3000
echo.
echo 📝 The Add-in will automatically detect and connect to the MCP server.
echo    Look for the 'LIVE' status in the Word task pane.
echo.
echo ⌨️  Press any key to stop both servers
echo.

pause >nul

REM Cleanup
echo.
echo 🛑 Shutting down servers...

REM Kill processes by port
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8765 "') do (
    if "%%a" neq "" taskkill /PID %%a /F >nul 2>&1
)

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":3000 "') do (
    if "%%a" neq "" taskkill /PID %%a /F >nul 2>&1
)

echo ✅ Cleanup complete
pause