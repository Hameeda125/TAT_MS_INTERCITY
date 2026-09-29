@echo off
cd /d "%~dp0"

echo ============================================================
echo TAT MS INTERCITY PIPELINE
echo ============================================================

echo.
echo Activating project virtual environment...
call "D:\Hameeda_Projects\TAT_MS_INTERCITY\.venv\Scripts\activate.bat"

echo.
echo Checking Python...
python --version

echo.
echo Checking PyTorch GPU...
python -c "import torch; print('PyTorch:', torch.__version__); print('CUDA available:', torch.cuda.is_available()); print('GPU:', torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'No CUDA GPU detected')"

echo.
echo ============================================================
echo Running pipeline...
echo ============================================================
python src\update_pipeline.py

echo.
echo ============================================================
echo Pipeline finished
echo ============================================================

pause