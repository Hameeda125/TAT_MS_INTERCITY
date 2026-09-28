@echo off
cd /d "%~dp0"

call "C:\Users\JAVED\anaconda3\Scripts\activate.bat"
call conda activate tat_pipeline

python src\update_pipeline.py

pause