@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo =========================================
echo        Office2PDF Build Script
echo =========================================
echo.

REM ����������ļ��Ƿ����
if not exist "Office2PDF.sln" (
    echo [ERROR] δ�ҵ� Office2PDF.sln �ļ�
    echo [ERROR] ��ȷ������Ŀ��Ŀ¼���д˽ű�
    pause
    exit /b 1
)

echo [INFO] �����������õ� Visual Studio �汾...
echo.

REM ����Visual Studio�汾��麯��
set "VS_FOUND=false"
set "VS_VERSION="
set "VS_PATH="

REM Visual Studio 2025
echo [CHECK] ��� Visual Studio 2025...
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
echo [CHECK] ��� Visual Studio 2022...
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
echo [CHECK] ��� Visual Studio 2019...
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
echo [CHECK] ��� Visual Studio 2017...
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

REM ���ɰ汾 Visual Studio (VS2015, VS2013, VS2012, VS2010)
echo [CHECK] ���ɰ汾 Visual Studio...
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
        REM ���ɰ汾���������ļ�·��
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

REM ��� MSBuild ������װ
echo [CHECK] �������� MSBuild...
for %%V in (Current 17.0 16.0 15.0 14.0) do (
    for %%A in ("C:\Program Files" "C:\Program Files (x86)") do (
        REM ��� VS2025 BuildTools
        set "MSBUILD_PATH=%%~A\Microsoft Visual Studio\2025\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
        if exist "!MSBUILD_PATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=MSBuild 2025 (�����汾)"
            goto build_with_msbuild
        )
        REM ��� VS2022 BuildTools
        set "MSBUILD_PATH=%%~A\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
        if exist "!MSBUILD_PATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=MSBuild 2022 (�����汾)"
            goto build_with_msbuild
        )
        REM ������ MSBuild ��װ
        set "MSBUILD_PATH=%%~A\MSBuild\%%V\Bin\MSBuild.exe"
        if exist "!MSBUILD_PATH!" (
            set "VS_FOUND=true"
            set "VS_VERSION=MSBuild %%V (�����汾)"
            goto build_with_msbuild
        )
    )
)

REM ���û���ҵ� Visual Studio������ .NET CLI
if "!VS_FOUND!"=="false" (
    echo [WARNING] δ�ҵ� Visual Studio������ʹ�� .NET CLI...
    goto dotnet_build
)

:found_vs
echo [SUCCESS] �ҵ�: !VS_VERSION!
echo [INFO] ·��: !VS_PATH!
echo [INFO] ���ڳ�ʼ�� Visual Studio ��������...
call "!VS_PATH!"
if errorlevel 1 (
    echo [ERROR] ��ʼ�� Visual Studio ����ʧ��
    goto dotnet_build
)
goto build_with_vs

:found_vs_legacy
echo [SUCCESS] �ҵ�: !VS_VERSION! (�ɰ汾)
echo [INFO] ·��: !VS_PATH!
echo [INFO] ���ڳ�ʼ�� Visual Studio ��������...
call "!VS_PATH!" x86
if errorlevel 1 (
    echo [ERROR] ��ʼ�� Visual Studio ����ʧ��
    goto dotnet_build
)
goto build_with_vs

:build_with_vs
echo.
echo [INFO] ʹ�� !VS_VERSION! ������Ŀ...
echo [INFO] ����: Release
echo [INFO] ƽ̨: Any CPU
echo.
msbuild Office2PDF.sln /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
if errorlevel 1 (
    echo.
    echo [ERROR] MSBuild ����ʧ��
    goto dotnet_build
)
goto build_success

:build_with_msbuild
echo [SUCCESS] �ҵ�: !VS_VERSION!
echo [INFO] ·��: !MSBUILD_PATH!
echo.
echo [INFO] ʹ�ö��� MSBuild ������Ŀ...
echo [INFO] ����: Release
echo [INFO] ƽ̨: Any CPU
echo.
"!MSBUILD_PATH!" Office2PDF.sln /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
if errorlevel 1 (
    echo.
    echo [ERROR] MSBuild ����ʧ��
    goto dotnet_build
)
goto build_success

:dotnet_build
echo.
echo [INFO] ����ʹ�� .NET CLI ����...
where dotnet >nul 2>nul
if errorlevel 1 (
    echo [ERROR] δ�ҵ� dotnet ����
    echo [ERROR] �밲װ .NET SDK �� Visual Studio
    echo.
    echo [HELP] ��������:
    echo [HELP] Visual Studio: https://visualstudio.microsoft.com/
    echo [HELP] .NET SDK: https://dotnet.microsoft.com/download
    goto build_failed
)

echo [INFO] ��� .NET �汾...
dotnet --version
echo.
echo [INFO] ʹ�� .NET CLI ������Ŀ...
echo [INFO] ����: Release
echo.
dotnet build Office2PDF.sln --configuration Release --verbosity minimal
if errorlevel 1 (
    echo.
    echo [ERROR] .NET CLI ����ʧ��
    goto build_failed
)
goto build_success

:build_success
echo.
echo =========================================
echo [SUCCESS] ����ɹ���ɣ�
echo =========================================
echo.
echo [INFO] ���Ŀ¼:
if exist "bin\Release" (
    echo [INFO] - bin\Release\
    dir /b "bin\Release\*.exe" 2>nul | findstr /r ".*" >nul && echo [INFO] ��ִ���ļ�������
)
if exist "bin\x86\Release" (
    echo [INFO] - bin\x86\Release\
)
if exist "bin\x64\Release" (
    echo [INFO] - bin\x64\Release\
)
echo.
echo [INFO] �������ʱ��: %date% %time%
@REM pause
exit /b 0

:build_failed
echo.
echo =========================================
echo [ERROR] ����ʧ�ܣ�
echo =========================================
echo.
echo [TROUBLESHOOT] �����ų�����:
echo [TROUBLESHOOT] 1. ȷ����װ�� Visual Studio �� .NET SDK
echo [TROUBLESHOOT] 2. ȷ����Ŀ�ļ�û����
echo [TROUBLESHOOT] 3. ����Ƿ�ȱ��������
echo [TROUBLESHOOT] 4. ������ Visual Studio ���ֶ�����
echo.
pause
exit /b 1

:end
