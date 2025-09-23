@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo 正在清理构建文件...

REM 删除编译输出目录
if exist "bin" (
    echo 删除 bin 目录...
    rmdir /s /q "bin" 2>nul
)

if exist "obj" (
    echo 删除 obj 目录...
    rmdir /s /q "obj" 2>nul
)

REM 删除 Visual Studio 缓存目录
if exist ".vs" (
    echo 删除 .vs 目录...
    rmdir /s /q ".vs" 2>nul
)

REM 删除 VSCode 缓存目录
if exist ".vscode" (
    echo 删除 .vscode 目录...
    rmdir /s /q ".vscode" 2>nul
)

REM 删除其他构建相关文件
for %%f in (*.user *.suo *.cache) do (
    if exist "%%f" (
        echo 删除 %%f...
        del /q "%%f" 2>nul
    )
)

REM 删除日志文件
for %%f in (*.log) do (
    if exist "%%f" (
        echo 删除 %%f...
        del /q "%%f" 2>nul
    )
)

REM 删除转换测试文件夹
if exist "test_PDFs" (
    echo 删除 test_PDFs 目录...
    rmdir /s /q "test_PDFs" 2>nul
)

REM 先删除目录下的所有空文件，再递归删除空文件夹（顺序不可颠倒，以免误删）
for /f "delims=" %%f in ('dir /a-d/b/s') do (
    if %%~zf==0 (
        del "%%f" 2>nul
    )
)
for /f "delims=" %%d in ('dir /ad/b/s') do (
    rd "%%d" 2>nul
)

echo 清理完成！
