@echo off

echo Cleaning old builds...

rmdir /s /q build
rmdir /s /q dist

del *.spec

echo Building EXE...

python -m PyInstaller --onefile --console --distpath release FF3PRPT_Comparison_Tool.py

echo.
echo Build complete!
pause