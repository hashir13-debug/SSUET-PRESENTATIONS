@echo off
title Animation Group 3 - Free Hosting Deployment
color 0B
echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║   ANIMATION GROUP 3 — FREE HOSTING DEPLOYER         ║
echo  ║   Sir Syed University — Graphics ^& Multimedia        ║
echo  ╚══════════════════════════════════════════════════════╝
echo.
echo  This will deploy your presentation to a FREE permanent link
echo  Example:  https://animation-group3.surge.sh
echo.
echo  Steps:
echo  1. Enter your email (creates free account)
echo  2. Enter a password
echo  3. Done! You get a permanent link instantly.
echo.
echo  ─────────────────────────────────────────────────────────
echo.
cd /d "%~dp0DEPLOY"
npx surge --domain animation-group3.surge.sh
echo.
echo  ─────────────────────────────────────────────────────────
echo  If the domain is taken, try: animation-grp3.surge.sh
echo  or run:  npx surge --domain YOUR-CHOICE.surge.sh
echo.
pause
