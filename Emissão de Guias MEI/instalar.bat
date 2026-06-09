@echo off
echo ============================================
echo  Instalando dependencias do Robo MEI...
echo ============================================

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m playwright install chromium

echo.
echo ============================================
echo  Instalacao concluida!
echo  Para rodar o robo: python main.py
echo ============================================
pause
