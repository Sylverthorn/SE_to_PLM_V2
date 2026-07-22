@echo off
echo ===================================================
echo   SE to PLM - Compilation en fichier EXE unique
echo ===================================================
echo.

:: Verification de l'environnement virtuel
if not exist "venv\Scripts\pyinstaller.exe" (
    echo [INFO] PyInstaller n'est pas installe dans le venv. Installation...
    call venv\Scripts\pip.exe install pyinstaller
)

echo [INFO] Lancement de la compilation avec PyInstaller...
echo.

venv\Scripts\pyinstaller.exe ^
    --onefile ^
    --windowed ^
    --icon="image-removebg-preview.ico" ^
    --add-data "SE_to_PLM/ui/resources;SE_to_PLM/ui/resources" ^
    --add-data "SE_to_PLM/ui/styles;SE_to_PLM/ui/styles" ^
    --name "SE_to_PLM" ^
    SE_to_PLM/app/main.py

if %ERRORLEVEL% equ 0 (
    echo.
    echo ===================================================
    echo [SUCCES] La compilation est terminee.
    echo L'executable se trouve dans le dossier 'dist' :
    echo dist\SE_to_PLM.exe
    echo ===================================================
) else (
    echo.
    echo [ERREUR] La compilation a echoue.
)
pause
