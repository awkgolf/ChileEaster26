@echo off
echo ⛏ Starting Geological Journal Build...
call npm run build
if errorlevel 1 (
  echo ❌ Build failed.
  exit /b 1
)
echo ✅ Build Complete!
pause
