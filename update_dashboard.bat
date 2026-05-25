@echo off
chcp 65001 >nul
title 集运业务看板 · 一键更新

echo.
echo  ╔══════════════════════════════════╗
echo  ║   集运业务订单看板 · 一键更新   ║
echo  ╚══════════════════════════════════╝
echo.
echo  前提：请确保桌面上已放置最新的
echo       【集运业务流量包数据统计.xlsx】
echo.

C:\Users\Faye\.workbuddy\binaries\python\envs\default\Scripts\python.exe C:\Users\Faye\WorkBuddy\2026-05-11-task-1\update_dashboard.py

if errorlevel 1 (
    echo.
    echo  ❌ 更新失败，请查看上方错误信息。
) else (
    echo.
    echo  💡 提示：双击 index.html 或访问 https://faye0366.github.io/order-dashboard/ 查看最新看板。
)

echo.
pause
