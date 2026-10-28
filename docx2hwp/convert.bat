@echo off
chcp 65001 >nul
if "%~1"=="" (
  python "%~dp0docx2hwp.py"
) else (
  python "%~dp0docx2hwp.py" %*
)
pause
