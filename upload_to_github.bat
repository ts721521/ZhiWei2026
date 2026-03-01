@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

:: 若终端 PATH 未刷新，优先使用常见安装路径
set "GIT_EXE="
where git >nul 2>&1 && set GIT_EXE=git
if not defined GIT_EXE (
    if exist "C:\Program Files\Git\cmd\git.exe" set "GIT_EXE=C:\Program Files\Git\cmd\git.exe"
)
if not defined GIT_EXE (
    echo [错误] 未找到 git。请先安装 Git 并重新打开终端。
    pause
    exit /b 1
)
if "%GIT_EXE%" neq "git" set "PATH=%PATH%;C:\Program Files\Git\cmd"

echo [1/4] 状态...
git status -s
echo.

echo [2/4] 添加所有更改（含新文件）...
git add -A
echo.

echo [3/4] 提交（请根据需要修改下方提交信息）...
set "MSG=Sync: code and docs %date% %time%"
if not "%~1"=="" set "MSG=%~1"
git commit -m "%MSG%"
if errorlevel 1 (
    echo 无变更或提交已取消。
    pause
    exit /b 0
)
echo.

echo [4/4] 推送到 GitHub...
git push -u origin main
if errorlevel 1 (
    echo.
    echo 若推送失败，请先登录 GitHub：
    echo   gh auth login
    echo 然后重新运行本脚本。
    pause
    exit /b 1
)

echo.
echo 已成功推送到 https://github.com/ts721521/ZhiWei2026
pause
