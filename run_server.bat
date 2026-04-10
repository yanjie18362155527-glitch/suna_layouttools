@echo off
cd /d "%~dp0"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
"D:\Anaconda\envs\suna_layout\python.exe" -m streamlit run my_project\main.py --server.port 8501 --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false
pause
