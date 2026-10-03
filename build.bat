@echo off
chcp 936 >nul 2>&1
cd /d "%~dp0"
title 编译 Clip2File

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
    echo [错误] 找不到 C# 编译器 csc.exe，请确认已安装 .NET Framework 4.x
    pause
    exit /b 1
)

echo 正在编译 Clip2File.exe ...
"%CSC%" /nologo /target:winexe /platform:anycpu /optimize+ /codepage:65001 ^
    /win32icon:build\app.ico ^
    /out:Clip2File.exe ^
    /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll ^
    src\Clip2File.cs

if errorlevel 1 (
    echo.
    echo [失败] 编译出错，请看上面的提示。
    pause
    exit /b 1
)

echo.
echo [成功] 已生成 Clip2File.exe
echo.
pause
