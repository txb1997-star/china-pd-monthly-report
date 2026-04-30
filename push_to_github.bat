@echo off
REM ============================================================
REM China PD Monthly Report - 一键 push 到 GitHub
REM ============================================================
REM 双击运行即可。会:
REM   1. 把最新 EN 版报表复制为 index.html(GitHub Pages 首页)
REM   2. git add 所有改动(.gitignore 会自动排除 API key / xlsx)
REM   3. 用当前时间戳 commit
REM   4. push 到 GitHub
REM
REM 第一次使用前必须先做过 git init / git remote add origin
REM (按 Claude 给的 step-by-step 命令做过一次)
REM ============================================================

setlocal enabledelayedexpansion

REM 切到本脚本所在目录
cd /d "%~dp0"

echo ============================================================
echo  China PD Monthly Report - Auto Push
echo ============================================================
echo.

REM ---- Step 1: 找到最新月份的 EN 版报表,复制为 index.html ----
echo [1/4] Updating index.html with latest EN report...

set "LATEST_EN="
for /f "delims=" %%F in ('dir /b /o-d "China_PD_Monthly_Report_*_EN.html" 2^>nul ^| findstr /v "_prev"') do (
    if not defined LATEST_EN set "LATEST_EN=%%F"
)

if not defined LATEST_EN (
    echo   [WARN] Cannot find any China_PD_Monthly_Report_*_EN.html
    echo   Skipping index.html update.
) else (
    copy /Y "!LATEST_EN!" "index.html" >nul
    echo   index.html ^<- !LATEST_EN!
)
echo.

REM ---- Step 2: git status ----
echo [2/4] Checking git status...
git status --short
echo.

REM ---- Step 3: add + commit ----
echo [3/4] Staging and committing...
git add .

REM 检查有没有改动可以 commit
git diff --cached --quiet
if !errorlevel! equ 0 (
    echo   No changes to commit. Skipping commit.
    echo.
    goto push_step
)

REM 用当前时间戳做 commit message
for /f "tokens=2 delims==" %%a in ('wmic os get LocalDateTime /value 2^>nul') do set "DT=%%a"
set "COMMIT_MSG=Update report - !DT:~0,4!-!DT:~4,2!-!DT:~6,2! !DT:~8,2!:!DT:~10,2!"

git commit -m "!COMMIT_MSG!"
if !errorlevel! neq 0 (
    echo   [ERROR] Commit failed.
    pause
    exit /b 1
)
echo.

:push_step
REM ---- Step 4: push ----
echo [4/4] Pushing to GitHub...
git push
if !errorlevel! neq 0 (
    echo.
    echo   [ERROR] Push failed.
    echo   Possible causes:
    echo     - 没装 git / 没配 remote(先按 step-by-step 命令跑一次)
    echo     - 网络问题
    echo     - 认证过期(浏览器会弹出 GitHub 登录,跟着登一下)
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Done. View your report at:
echo  https://txb1997-star.github.io/china-pd-monthly-report/
echo ============================================================
echo.
pause
endlocal
