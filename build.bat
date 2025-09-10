@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo =========================================
echo        Office2PDF Build Script
echo =========================================
echo.

REM 检查解决方案文件是否存在
if not exist "Office2PDF.sln" (
    echo [ERROR] 未找到 Office2PDF.sln 文件
    echo [ERROR] 请确保在项目根目录运行此脚本
    pause
    exit /b 1
)

echo [INFO] 正在搜索可用的 Visual Studio 版本...
echo.

REM 定义Visual Studio版本检查函数
set "VS_FOUND=false"
set "VS_VERSION="
set "VS_PATH="

REM Visual Studio 2025
echo [CHECK] 检查 Visual Studio 2025...
for %%E in (Preview Enterprise Professional Community BuildTools) do (
    for %%A in ("C:\Program Files" "C:\Program Files (x86)") do (
        set "VSPATH=%%~A\Microsoft Visual Studio\2025\%%E\Common7\Tools\VsDevCmd.bat"
        if exist "!VSPATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=Visual Studio 2025 %%E"
            set "VS_PATH=!VSPATH!"
            goto found_vs
        )
    )
)

REM Visual Studio 2022
echo [CHECK] 检查 Visual Studio 2022...
for %%E in (Preview Enterprise Professional Community BuildTools) do (
    for %%A in ("C:\Program Files" "C:\Program Files (x86)") do (
        set "VSPATH=%%~A\Microsoft Visual Studio\2022\%%E\Common7\Tools\VsDevCmd.bat"
        if exist "!VSPATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=Visual Studio 2022 %%E"
            set "VS_PATH=!VSPATH!"
            goto found_vs
        )
    )
)

REM Visual Studio 2019
echo [CHECK] 检查 Visual Studio 2019...
for %%E in (Preview Enterprise Professional Community BuildTools) do (
    for %%A in ("C:\Program Files" "C:\Program Files (x86)") do (
        set "VSPATH=%%~A\Microsoft Visual Studio\2019\%%E\Common7\Tools\VsDevCmd.bat"
        if exist "!VSPATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=Visual Studio 2019 %%E"
            set "VS_PATH=!VSPATH!"
            goto found_vs
        )
    )
)

REM Visual Studio 2017
echo [CHECK] 检查 Visual Studio 2017...
for %%E in (Enterprise Professional Community BuildTools) do (
    for %%A in ("C:\Program Files" "C:\Program Files (x86)") do (
        set "VSPATH=%%~A\Microsoft Visual Studio\2017\%%E\Common7\Tools\VsDevCmd.bat"
        if exist "!VSPATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=Visual Studio 2017 %%E"
            set "VS_PATH=!VSPATH!"
            goto found_vs
        )
    )
)

REM 检查旧版本 Visual Studio (VS2015, VS2013, VS2012, VS2010)
echo [CHECK] 检查旧版本 Visual Studio...
for %%V in (14.0 12.0 11.0 10.0) do (
    for %%A in ("C:\Program Files" "C:\Program Files (x86)") do (
        set "VSPATH=%%~A\Microsoft Visual Studio %%V\Common7\Tools\VsDevCmd.bat"
        if exist "!VSPATH!" (
            set "VS_FOUND=true"
            if "%%V"=="14.0" set "VS_VERSION=Visual Studio 2015"
            if "%%V"=="12.0" set "VS_VERSION=Visual Studio 2013"
            if "%%V"=="11.0" set "VS_VERSION=Visual Studio 2012"
            if "%%V"=="10.0" set "VS_VERSION=Visual Studio 2010"
            set "VS_PATH=!VSPATH!"
            goto found_vs
        )
        REM 检查旧版本的批处理文件路径
        set "VSPATH=%%~A\Microsoft Visual Studio %%V\VC\vcvarsall.bat"
        if exist "!VSPATH!" (
            set "VS_FOUND=true"
            if "%%V"=="14.0" set "VS_VERSION=Visual Studio 2015"
            if "%%V"=="12.0" set "VS_VERSION=Visual Studio 2013"
            if "%%V"=="11.0" set "VS_VERSION=Visual Studio 2012"
            if "%%V"=="10.0" set "VS_VERSION=Visual Studio 2010"
            set "VS_PATH=!VSPATH!"
            goto found_vs_legacy
        )
    )
)

REM 检查 MSBuild 独立安装
echo [CHECK] 检查独立的 MSBuild...
for %%V in (Current 17.0 16.0 15.0 14.0) do (
    for %%A in ("C:\Program Files" "C:\Program Files (x86)") do (
        REM 检查 VS2025 BuildTools
        set "MSBUILD_PATH=%%~A\Microsoft Visual Studio\2025\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
        if exist "!MSBUILD_PATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=MSBuild 2025 (独立版本)"
            goto build_with_msbuild
        )
        REM 检查 VS2022 BuildTools
        set "MSBUILD_PATH=%%~A\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
        if exist "!MSBUILD_PATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=MSBuild 2022 (独立版本)"
            goto build_with_msbuild
        )
        REM 检查独立 MSBuild 安装
        set "MSBUILD_PATH=%%~A\MSBuild\%%V\Bin\MSBuild.exe"
        if exist "!MSBUILD_PATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=MSBuild %%V (独立版本)"
            goto build_with_msbuild
        )
    )
)

REM 如果没有找到 Visual Studio，尝试 .NET CLI
if "!VS_FOUND!"=="false" (
    echo [WARNING] 未找到 Visual Studio，尝试使用 .NET CLI...
    goto dotnet_build
)

:found_vs
echo [SUCCESS] 找到: !VS_VERSION!
echo [INFO] 路径: !VS_PATH!
echo [INFO] 正在初始化 Visual Studio 开发环境...
call "!VS_PATH!"
if errorlevel 1 (
    echo [ERROR] 初始化 Visual Studio 环境失败
    goto dotnet_build
)
goto build_with_vs

:found_vs_legacy
echo [SUCCESS] 找到: !VS_VERSION! (旧版本)
echo [INFO] 路径: !VS_PATH!
echo [INFO] 正在初始化 Visual Studio 开发环境...
call "!VS_PATH!" x86
if errorlevel 1 (
    echo [ERROR] 初始化 Visual Studio 环境失败
    goto dotnet_build
)
goto build_with_vs

:build_with_vs
echo.
echo [INFO] 使用 !VS_VERSION! 编译项目...
echo [INFO] 配置: Release
echo [INFO] 平台: Any CPU
echo.
msbuild Office2PDF.sln /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
if errorlevel 1 (
    echo.
    echo [ERROR] MSBuild 编译失败
    goto dotnet_build
)
goto build_success

:build_with_msbuild
echo [SUCCESS] 找到: !VS_VERSION!
echo [INFO] 路径: !MSBUILD_PATH!
echo.
echo [INFO] 使用独立 MSBuild 编译项目...
echo [INFO] 配置: Release
echo [INFO] 平台: Any CPU
echo.
"!MSBUILD_PATH!" Office2PDF.sln /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
if errorlevel 1 (
    echo.
    echo [ERROR] MSBuild 编译失败
    goto dotnet_build
)
goto build_success

:dotnet_build
echo.
echo [INFO] 尝试使用 .NET CLI 编译...
where dotnet >nul 2>nul
if errorlevel 1 (
    echo [ERROR] 未找到 dotnet 命令
    echo [ERROR] 请安装 .NET SDK 或 Visual Studio
    echo.
    echo [HELP] 下载链接:
    echo [HELP] Visual Studio: https://visualstudio.microsoft.com/
    echo [HELP] .NET SDK: https://dotnet.microsoft.com/download
    goto build_failed
)

echo [INFO] 检查 .NET 版本...
dotnet --version
echo.
echo [INFO] 使用 .NET CLI 编译项目...
echo [INFO] 配置: Release
echo.
dotnet build Office2PDF.sln --configuration Release --verbosity minimal
if errorlevel 1 (
    echo.
    echo [ERROR] .NET CLI 编译失败
    goto build_failed
)
goto build_success

:build_success
echo.
echo =========================================
echo [SUCCESS] 编译成功完成！
echo =========================================
echo.
echo [INFO] 输出目录:
if exist "bin\Release" (
    echo [INFO] - bin\Release\
    dir /b "bin\Release\*.exe" 2>nul | findstr /r ".*" >nul && echo [INFO] 可执行文件已生成
)
if exist "bin\x86\Release" (
    echo [INFO] - bin\x86\Release\
)
if exist "bin\x64\Release" (
    echo [INFO] - bin\x64\Release\
)
echo.
echo [INFO] 编译完成时间: %date% %time%
@REM pause
exit /b 0

:build_failed
echo.
echo =========================================
echo [ERROR] 编译失败！
echo =========================================
echo.
echo [TROUBLESHOOT] 故障排除建议:
echo [TROUBLESHOOT] 1. 确保安装了 Visual Studio 或 .NET SDK
echo [TROUBLESHOOT] 2. 确保项目文件没有损坏
echo [TROUBLESHOOT] 3. 检查是否缺少依赖项
echo [TROUBLESHOOT] 4. 尝试在 Visual Studio 中手动编译
echo.
pause
exit /b 1

:end
