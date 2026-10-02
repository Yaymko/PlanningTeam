@echo off
rem Запуск на общей машине: приложение доступно коллегам по сети на http://<имя-машины>:8501
cd /d "%~dp0"
python -m pip install -r requirements.txt --quiet
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501 --server.headless true
pause
