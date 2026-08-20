@echo off
cd /d "%~dp0"
title Vimar Datasheet Tool
echo Starting the Vimar Datasheet Tool...
echo Your browser will open automatically in a few seconds.
echo.
echo Keep this window open while you use the tool.
echo Close this window when you are done.
echo.
python -m streamlit run app.py
pause
