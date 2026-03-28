@echo off
chcp 65001 > nul
echo Web版を起動しています... (Starting Web App...)
echo ブラウザが自動的に開きます。閉じないでください。
python -m streamlit run src/streamlit_app.py
pause
