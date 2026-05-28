@echo off
chcp 65001 >nul
title 集运业务看板 · 一键更新并推送

echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║   集运业务订单看板 · 一键更新并推送到 GitHub ║
echo  ╚══════════════════════════════════════════════╝
echo.
echo  前提：请确保桌面 WorkBuddy数据看板 文件夹中
echo       已放置最新的【集运业务流量包数据统计.xlsx】
echo.

:: Step 1: 更新看板数据
C:\Users\Faye\.workbuddy\binaries\python\envs\default\Scripts\python.exe C:\Users\Faye\WorkBuddy\2026-05-11-task-1\update_dashboard.py

if errorlevel 1 (
    echo.
    echo  ❌ 数据更新失败，请查看上方错误信息。
    echo.
    pause
    exit /b 1
)

:: Step 2: 推送到 GitHub
echo.
echo  🚀 正在推送到 GitHub...
echo.
cd /d C:\Users\Faye\WorkBuddy\2026-05-11-task-1
"C:\Program Files\Git\bin\git.exe" add .
"C:\Program Files\Git\bin\git.exe" commit -m "update: dashboard data %date% %time%"
"C:\Program Files\Git\bin\git.exe" push origin main

if errorlevel 1 (
    echo.
    echo  ❌ GitHub 推送失败，请检查网络连接。
) else (
    echo.
    echo  ✅ 更新完成！
    echo.
    echo  📎 分享链接：https://faye0366.github.io/order-dashboard/
    echo     （同事刷新链接即可查看最新数据）
)

echo.
pause
