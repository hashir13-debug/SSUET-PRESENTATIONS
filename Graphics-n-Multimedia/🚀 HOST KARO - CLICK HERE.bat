@echo off
title Presentation Host Karo - Free Link
color 0A
cls
echo.
echo  ╔══════════════════════════════════════════════════════════════╗
echo  ║     🚀  ANIMATION GROUP 3 — FREE HOSTING                    ║
echo  ║     Sir Syed University of Engineering ^& Technology          ║
echo  ╚══════════════════════════════════════════════════════════════╝
echo.
echo  ✅  2 ASAAN OPTION HAIN - DONO FREE HAI HAMESHA KE LIYE
echo.
echo  ══════════════════════════════════════════════════════════════
echo  OPTION 1:  NETLIFY DROP  (Sirf Drag ^& Drop — 30 seconds)
echo  ══════════════════════════════════════════════════════════════
echo.
echo   Step 1: Browser khulega → netlify.com/drop
echo   Step 2: File Explorer khulega → DEPLOY folder dikhega
echo   Step 3: DEPLOY folder ko browser mein drag karein
echo   Step 4: Done! Link milega jaise: abc123.netlify.app
echo.
echo  ══════════════════════════════════════════════════════════════
echo  OPTION 2:  SURGE.SH  (Command se — 1 minute)
echo  ══════════════════════════════════════════════════════════════
echo.
echo   Step 1: Koi bhi email daalein (naya free account banta hai)
echo   Step 2: Password daalein (apna banao)
echo   Step 3: Link milega:  animation-group3-ssuet.surge.sh
echo.
echo  ══════════════════════════════════════════════════════════════
echo.
echo  Konsa option chahiye? (1 ya 2 daalen)
echo.
set /p choice="  Apni choice: "
echo.

if "%choice%"=="1" goto netlify
if "%choice%"=="2" goto surge
goto netlify

:netlify
echo  ▶ Netlify Drop khol raha hoon...
echo.
start "" "https://app.netlify.com/drop"
timeout /t 2 /nobreak >nul
echo  ▶ DEPLOY folder khol raha hoon...
start "" "explorer.exe" "%~dp0DEPLOY"
echo.
echo  ────────────────────────────────────────────────────────────
echo  ✅ Dono window khul gaye!
echo.
echo  Ab karo:
echo  1. DEPLOY folder mein se "DEPLOY" folder ko
echo     Netlify browser window mein drag karein
echo  2. Account banao (free hai)
echo  3. Link copy karein — done!
echo  ────────────────────────────────────────────────────────────
goto end

:surge
echo  ▶ Surge.sh deploy shuru ho raha hai...
echo.
echo  ⚠  Pehla field mein:  Koi bhi email daalein (apna banao)
echo  ⚠  Doosra field mein: Password banao (yaad rakhna)
echo  ✅ Link milega: animation-group3-ssuet.surge.sh
echo.
cd /d "%~dp0DEPLOY"
npx surge . animation-group3-ssuet.surge.sh
goto end

:end
echo.
echo  ════════════════════════════════════════════════════════════
echo  ✅ Hosting complete! Link share karo classmates ke saath 🎉
echo  ════════════════════════════════════════════════════════════
echo.
pause
