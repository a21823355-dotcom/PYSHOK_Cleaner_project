@echo off
REM ============================================================
REM  Сборка PYSHOK_Cleaner.exe из исходника cleaner.py
REM  Запускать на Windows, двойным кликом, из этой же папки.
REM  Требуется установленный Python (python.org, галочка "Add to PATH").
REM ============================================================

echo Проверка Python...
python --version
if errorlevel 1 (
    echo.
    echo Python не найден. Установите Python с https://python.org
    echo При установке ОБЯЗАТЕЛЬНО поставьте галочку "Add python.exe to PATH".
    pause
    exit /b 1
)

echo.
echo Установка PyInstaller (если ещё не установлен)...
python -m pip install --upgrade pyinstaller

echo.
echo Сборка EXE...
python -m PyInstaller --onefile --windowed ^
    --name "PYSHOK_Cleaner" ^
    --version-file "version_info.txt" ^
    cleaner.py

echo.
if exist "dist\PYSHOK_Cleaner.exe" (
    echo Готово! Файл находится в папке dist\PYSHOK_Cleaner.exe
) else (
    echo Что-то пошло не так, проверьте вывод выше на ошибки.
)
pause
