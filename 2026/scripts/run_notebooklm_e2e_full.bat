@echo off
REM 全量 NotebookLM E2E（6458 文件），在本机终端运行，输出重定向到日志。
REM 建议：先确保 Z:\Schneider\5_投标 与 D:\ZWPDFTSEST 可访问。
cd /d "%~dp0.."
set LOG=docs\test-reports\notebooklm_e2e_full_%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log
set LOG=%LOG: =0%
echo [E2E Full] Starting at %date% %time%, log: %LOG%
python scripts/run_notebooklm_e2e.py > "%LOG%" 2>&1
echo [E2E Full] Exit code: %ERRORLEVEL%, see %LOG%
exit /b %ERRORLEVEL%
