@echo off
chcp 65001 >nul
rem ==========================================================================
rem  Gera o instalador no proprio PC (opcional - o GitHub Actions ja faz isso).
rem  Requisitos: Python 3.12+, Inno Setup 6 e Tesseract instalado em
rem  C:\Program Files\Tesseract-OCR com o idioma Portuguese.
rem ==========================================================================
cd /d "%~dp0"

for /f "usebackq" %%v in (`python -c "import sys; sys.path.insert(0,'src'); import versao; print(versao.VERSAO)"`) do set VERSAO=%%v
if "%VERSAO%"=="" goto erro
echo Versao: %VERSAO%

echo.
echo [1/3] Bibliotecas e executavel...
python -m pip install -r requirements.txt "pyinstaller>=6,<7" || goto erro
python -m PyInstaller --noconfirm --clean --windowed --name BuscaDatabook --paths src ^
  --collect-all pypdfium2 --collect-all docx --collect-all pptx ^
  --hidden-import win32com.client src\app.py || goto erro

echo.
echo [2/3] Copiando o Tesseract...
if not exist "C:\Program Files\Tesseract-OCR\tessdata\por.traineddata" (
  echo Tesseract com idioma Portuguese nao encontrado em C:\Program Files\Tesseract-OCR
  goto erro
)
robocopy "C:\Program Files\Tesseract-OCR" vendor\tesseract /E /XF unins*.* /NFL /NDL /NJH /NJS >nul
if errorlevel 8 goto erro

echo.
echo [3/3] Instalador...
"%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" /DAppVersion=%VERSAO% installer\BuscaDatabook.iss || goto erro
powershell -NoProfile -Command "(Get-FileHash 'output\BuscaDatabook-Setup-%VERSAO%.exe' -Algorithm SHA256).Hash.ToLower()"
echo.
echo Pronto: output\BuscaDatabook-Setup-%VERSAO%.exe
explorer output
pause
exit /b 0

:erro
echo.
echo Algo deu errado. Veja a mensagem acima.
pause
exit /b 1