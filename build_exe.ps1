python -m venv .venv
.venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller==6.10.0

pyinstaller --clean -y build.spec

Write-Host "Built: dist\mplus2excel.exe"
