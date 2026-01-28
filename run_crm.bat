@echo off
echo ================================================
echo    Mise a jour SUIVI_CRM
echo ================================================
echo.

REM Changer vers le repertoire du script
cd /d "%~dp0"

REM Activer l'environnement virtuel si present
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

REM Lancer le script Python
python scripts\update_crm.py

REM Pause pour voir les messages
pause
