@echo off
cd /d %~dp0
python build_dashboard.py --out docs/index.html
if errorlevel 1 goto end
git add docs/index.html config
git commit -m "Update dashboard %date%"
git push
:end
pause
