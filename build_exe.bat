@echo off
echo ================================================
echo    Build de l'executable
echo ================================================
echo.

REM Verifier que PyInstaller est installe
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installation de PyInstaller...
    pip install pyinstaller
)

REM Creer l'executable
echo.
echo Creation de l'executable...
pyinstaller --onedir --name "Automatisation_SUIVI" --console ^
    --hidden-import=scripts.update_kpis ^
    --hidden-import=scripts.update_mdr ^
    --hidden-import=scripts.update_pma ^
    --hidden-import=scripts.update_produit ^
    --hidden-import=scripts.update_crm ^
    --hidden-import=config ^
    --hidden-import=src.excel_automation ^
    --add-data "config.py;." ^
    --add-data "scripts;scripts" ^
    --add-data "src;src" ^
    launcher.py

echo.
echo ================================================
echo    BUILD TERMINE
echo ================================================
echo.
echo Le dossier se trouve dans: dist\Automatisation_SUIVI\
echo.
echo Pour envoyer au client:
echo 1. Compressez le dossier dist\Automatisation_SUIVI en ZIP
echo 2. Envoyez le ZIP au client
echo 3. Le client dezippe et double-clique sur Automatisation_SUIVI.exe
echo.
pause
