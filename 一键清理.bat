@echo off
setlocal enabledelayedexpansion

REM ����Ƿ��� UNC ����·����
set "SCRIPT_DIR=%~dp0"
echo �ű�Ŀ¼: %SCRIPT_DIR%

REM �����л����ű�����Ŀ¼
REM ����� UNC ·����cd /d ��ʧ�ܣ������ǿ���ʹ�� pushd ��ӳ����ʱ������
if "%SCRIPT_DIR:~0,2%"=="\\" (
    echo ��⵽ UNC ����·��������ӳ����ʱ������...
    pushd "%SCRIPT_DIR%" || (
        echo ����: �޷����� UNC ·����������������
        pause
        exit /b 1
    )
) else (
    cd /d "%SCRIPT_DIR%" || (
        echo ����: �޷��л����ű�Ŀ¼
        pause
        exit /b 1
    )
)

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

REM ɾ��ת�������ļ���
if exist "test_PDFs" (
    echo ɾ�� test_PDFs Ŀ¼...
    rmdir /s /q "test_PDFs" 2>nul
)

REM ��ɾ��Ŀ¼�µ����п��ļ����ٵݹ�ɾ�����ļ��У�˳�򲻿ɵߵ���������ɾ��
for /f "delims=" %%f in ('dir /a-d/b/s') do (
    if %%~zf==0 (
        del "%%f" 2>nul
    )
)
for /f "delims=" %%d in ('dir /ad/b/s') do (
    rd "%%d" 2>nul
)

echo ������ɣ�

REM ����� UNC ·������Ҫʹ�� popd ��������ʱӳ��
if "%SCRIPT_DIR:~0,2%"=="\\" (
    popd
)

