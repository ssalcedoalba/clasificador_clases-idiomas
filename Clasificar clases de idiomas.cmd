@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Clasificador de clases de idiomas

echo ==================================================
echo          CLASIFICADOR DE CLASES DE IDIOMAS
echo ==================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: No se encontro el entorno virtual .venv.
    echo Revisa la instalacion antes de continuar.
    echo.
    pause
    exit /b 2
)

if not exist "clasificar_clases_idiomas.py" (
    echo ERROR: No se encontro clasificar_clases_idiomas.py.
    echo.
    pause
    exit /b 2
)

echo Revisando la carpeta entrada...
echo.
".venv\Scripts\python.exe" "clasificar_clases_idiomas.py"
set "CODIGO=%ERRORLEVEL%"

echo.
if "%CODIGO%"=="0" (
    echo Proceso terminado correctamente.
    if exist "salida\clases.xlsx" (
        echo Se abrira el Excel consolidado.
        start "" "%~dp0salida\clases.xlsx"
    ) else (
        echo Se abrira la carpeta de resultados.
        if exist "salida" start "" "%~dp0salida"
    )
) else (
    echo El programa termino con uno o mas errores.
    echo Lee el mensaje anterior para conocer la causa.
)

echo.
echo Presiona una tecla para cerrar esta ventana.
pause >nul
exit /b %CODIGO%
