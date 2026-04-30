@echo off
REM ================================================================
REM  China PD Monthly Report -- One-Click Push
REM ================================================================
REM  双击我!我会自动:
REM    1. 检查 git 是否安装(没装就告诉你去哪下)
REM    2. 第一次跑:清理 / 初始化 git repo,连接到 GitHub,强推一次
REM    3. 之后跑:把最新 EN 报表复制成 index.html,然后正常 push
REM
REM  GitHub repo: https://github.com/txb1997-star/china-pd-monthly-report
REM  Pages URL  : https://txb1997-star.github.io/china-pd-monthly-report/
REM
REM  第一次跑会弹浏览器让你登 GitHub(用 txb1997-star 那个账号),
REM  Git 会帮你记住,以后不会再弹。
REM ================================================================

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ================================================================
echo  China PD Monthly Report  -  One-Click Push
echo ================================================================
echo.

REM ---- 0. 检查 git 是否安装 ----
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git is not installed.
    echo.
    echo   Please download and install Git for Windows from:
    echo   https://git-scm.com/download/win
    echo.
    echo   Then double-click this file again.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('git --version') do echo   %%v detected.
echo.

REM ---- 1. 检查 .git 状态 ----
set "FIRST_PUSH=0"
set "REPO_HEALTHY=0"

if exist ".git\config" (
    git rev-parse --is-inside-work-tree >nul 2>&1
    if not errorlevel 1 set "REPO_HEALTHY=1"
)

REM 如果 .git 存在但坏了(沙箱创建的半截),清掉
if exist ".git" if "!REPO_HEALTHY!"=="0" (
    echo [SETUP] Cleaning up broken .git folder from previous attempt...
    attrib -h -s -r ".git" /s /d >nul 2>&1
    rmdir /s /q ".git" 2>nul
    if exist ".git" (
        echo   [WARN] Could not auto-delete .git. Trying with takeown...
        takeown /f ".git" /r /d Y >nul 2>&1
        icacls ".git" /grant "%USERNAME%:F" /t /q >nul 2>&1
        rmdir /s /q ".git" 2>nul
    )
    if exist ".git" (
        echo   [ERROR] Cannot delete .git folder. Please manually delete it:
        echo     1. Open File Explorer
        echo     2. Click "View" -^> Show -^> Hidden items
        echo     3. Delete the .git folder in this directory
        echo     4. Run this script again
        pause
        exit /b 1
    )
    echo   Cleaned.
    echo.
)

REM ---- 2. 如果还没初始化,做首次 setup ----
if not exist ".git" (
    echo [SETUP] First-time setup: initializing git repo...
    git init -b main
    git config user.name "Summer Tan"
    git config user.email "xtan@chefman.com"
    git remote add origin https://github.com/txb1997-star/china-pd-monthly-report.git
    REM 关闭 OneDrive 容易出问题的 CRLF auto-conversion
    git config core.autocrlf false
    git config core.safecrlf false
    set "FIRST_PUSH=1"
    echo   Repo initialized and linked to GitHub.
    echo.
)

REM ---- 3. 更新 index.html: 用最新的 EN 报表 ----
echo [UPDATE] Refreshing index.html with latest EN report...
set "LATEST_EN="
for /f "delims=" %%F in ('dir /b /o-d "China_PD_Monthly_Report_*_EN.html" 2^>nul ^| findstr /v "_prev"') do (
    if not defined LATEST_EN set "LATEST_EN=%%F"
)
if defined LATEST_EN (
    copy /Y "!LATEST_EN!" "index.html" >nul
    echo   index.html  ^<--  !LATEST_EN!
) else (
    echo   [WARN] No China_PD_Monthly_Report_*_EN.html found, skipping.
)
echo.

REM ---- 4. 看看 git status ----
echo [CHECK] Files to be tracked:
git status --short
echo.

REM ---- 5. add ----
echo [ADD] Staging files...
git add .
echo.

REM ---- 6. 验证敏感文件确实没被加进去 ----
git diff --cached --name-only > "%TEMP%\_staged_files.txt"
findstr /i "claude_api_key.txt" "%TEMP%\_staged_files.txt" >nul && (
    echo [ABORT] claude_api_key.txt is staged! .gitignore may be wrong.
    echo Aborting before pushing secrets to a public repo.
    del "%TEMP%\_staged_files.txt" >nul 2>&1
    pause
    exit /b 1
)
findstr /i "\.xlsx" "%TEMP%\_staged_files.txt" >nul && (
    echo [ABORT] An xlsx file is staged! .gitignore may be wrong.
    echo Aborting before pushing internal data to a public repo.
    del "%TEMP%\_staged_files.txt" >nul 2>&1
    pause
    exit /b 1
)
del "%TEMP%\_staged_files.txt" >nul 2>&1
echo [SAFETY] No API key or xlsx in staged files. Good.
echo.

REM ---- 7. commit ----
git diff --cached --quiet
if !errorlevel! equ 0 (
    echo [COMMIT] Nothing to commit, working tree clean.
    echo.
    goto push_step
)

if "!FIRST_PUSH!"=="1" (
    set "COMMIT_MSG=Initial: build system + latest monthly report"
) else (
    for /f "tokens=2 delims==" %%a in ('wmic os get LocalDateTime /value 2^>nul') do set "DT=%%a"
    set "COMMIT_MSG=Update report - !DT:~0,4!-!DT:~4,2!-!DT:~6,2! !DT:~8,2!:!DT:~10,2!"
)

echo [COMMIT] !COMMIT_MSG!
git commit -m "!COMMIT_MSG!"
if !errorlevel! neq 0 (
    echo [ERROR] Commit failed.
    pause
    exit /b 1
)
echo.

:push_step
REM ---- 8. push ----
echo [PUSH] Pushing to GitHub...
if "!FIRST_PUSH!"=="1" (
    echo   First push - using --force to overwrite remote placeholder commits.
    echo   A browser window may open for GitHub login. Please sign in.
    echo.
    git push -u origin main --force
) else (
    git push
)

if !errorlevel! neq 0 (
    echo.
    echo [ERROR] Push failed.
    echo Possible causes:
    echo   - Authentication declined / cancelled in browser
    echo   - Network problem
    echo   - Repo settings changed on GitHub side
    pause
    exit /b 1
)

echo.
echo ================================================================
echo  DONE.
echo.
echo  Code:  https://github.com/txb1997-star/china-pd-monthly-report
echo  Live:  https://txb1997-star.github.io/china-pd-monthly-report/
echo         (GitHub Pages takes 1-2 min to rebuild)
echo ================================================================
echo.
pause
endlocal
