@echo off
title Push HALO Dashboard to GitHub
cd /d "%~dp0"

echo Pushing to GitHub...
git add -A
git commit -m "Update dashboard data %date% %time%"
git push origin main

echo.
echo Done! Dashboard will update on GitHub Pages in ~1 minute.
pause
