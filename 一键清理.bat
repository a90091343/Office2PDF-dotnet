@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo �����������ļ�...

REM ɾ���������Ŀ¼
if exist "bin" (
    echo ɾ�� bin Ŀ¼...
    rmdir /s /q "bin" 2>nul
)

if exist "obj" (
    echo ɾ�� obj Ŀ¼...
    rmdir /s /q "obj" 2>nul
)

REM ɾ�� Visual Studio ����Ŀ¼
if exist ".vs" (
    echo ɾ�� .vs Ŀ¼...
    rmdir /s /q ".vs" 2>nul
)

REM ɾ�� VSCode ����Ŀ¼
if exist ".vscode" (
    echo ɾ�� .vscode Ŀ¼...
    rmdir /s /q ".vscode" 2>nul
)

REM ɾ��������������ļ�
for %%f in (*.user *.suo *.cache) do (
    if exist "%%f" (
        echo ɾ�� %%f...
        del /q "%%f" 2>nul
    )
)

REM ɾ����־�ļ�
for %%f in (*.log) do (
    if exist "%%f" (
        echo ɾ�� %%f...
        del /q "%%f" 2>nul
    )
)

echo ������ɣ�
