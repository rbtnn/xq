@echo off
setlocal

cd /d "%~dp0"

where dotnet >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] dotnet not found. Install .NET SDK 10.
    echo         winget install Microsoft.DotNet.SDK.10
    exit /b 1
)

echo build: dotnet publish -c Release -r win-x64
dotnet publish "%~dp0xq.csproj" -c Release -r win-x64 ^
    --self-contained -p:PublishSingleFile=true -o "%~dp0dist"
if errorlevel 1 (
    echo.
    echo [ERROR] build failed.
    exit /b 1
)

if exist "%~dp0dist\xq.exe" (
    copy /y "%~dp0dist\xq.exe" "%~dp0xq.exe" >nul
)

echo.
echo built: xq.exe
endlocal
