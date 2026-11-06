@echo off
rem ---------------------------------------------------------------------------
rem  Praesentationsassistent starten (Testphase, lokal). Angelegt/aufgerufen
rem  von der Desktop-Verknuepfung bzw. einrichten.ps1.
rem
rem  "python -P": den aktuellen Ordner NICHT in den Suchpfad aufnehmen. Sonst
rem  wuerde beim Start aus dem Projektordner der dort liegende Quellcode-
rem  Ordner statt der in .venv installierten Programmversion geladen.
rem ---------------------------------------------------------------------------
setlocal enabledelayedexpansion
for %%I in ("%~dp0..") do set "ROOT=%%~fI"
set "PY=%ROOT%\.venv\Scripts\python.exe"
set "AUSGABE=%ROOT%\ausgabe"
if not exist "%PY%" (
  echo FEHLER: %PY% fehlt. Bitte windows\einrichten.ps1 erst ausfuehren.
  pause
  exit /b 1
)
if not exist "%AUSGABE%" mkdir "%AUSGABE%"

rem praesentation.env lesen (NAME=wert pro Zeile, # = Kommentar)
if exist "%ROOT%\praesentation.env" (
  for /f "usebackq eol=# tokens=* delims=" %%L in ("%ROOT%\praesentation.env") do set "%%L"
)

:menu
echo.
echo ================= Praesentationsassistent =================
echo  1. Test-Praesentation erzeugen (TESTDATEN, zum Ausprobieren)
echo  2. Beispiel mit echter Marke erzeugen (braucht Langdock-Key)
echo  3. Kommentare aus einer vorhandenen .pptx anzeigen
echo  4. Ueberarbeitungs-Oberflaeche im Browser oeffnen
echo  0. Beenden
echo ==============================================================
set "WAHL="
set /p "WAHL=Auswahl: "

if "%WAHL%"=="1" goto testpraesentation
if "%WAHL%"=="2" goto markenbeispiel
if "%WAHL%"=="3" goto kommentare
if "%WAHL%"=="4" goto oberflaeche
if "%WAHL%"=="0" goto ende
echo Ungueltige Auswahl.
goto menu

:testpraesentation
set "ZIEL=%AUSGABE%\test_praesentation.pptx"
"%PY%" -P -m praesentationsassistent.cli erzeugen --ausgabe "%ZIEL%"
if exist "%ZIEL%" echo Erzeugt: %ZIEL%
pause
goto menu

:markenbeispiel
if "%LANGDOCK_API_KEY%"=="" (
  echo FEHLER: In praesentation.env ist noch kein LANGDOCK_API_KEY eingetragen.
  pause
  goto menu
)
set "MARKE="
set "THEMA="
echo Marken: AMERON / URBAN LOFT / Althoff Hotels / Althoff Collection
set /p "MARKE=Marke: "
set /p "THEMA=Thema der Praesentation: "
if "%MARKE%"=="" goto menu
if "%THEMA%"=="" goto menu
set "ZIEL=%AUSGABE%\beispiel_%RANDOM%.pptx"
"%PY%" -P -m praesentationsassistent.cli erzeugen-mit-marke "%MARKE%" "%THEMA%" --ausgabe "%ZIEL%"
if exist "%ZIEL%" echo Erzeugt: %ZIEL%
pause
goto menu

:kommentare
set "DATEI="
set /p "DATEI=Pfad zur .pptx-Datei: "
if "%DATEI%"=="" goto menu
"%PY%" -P -m praesentationsassistent.cli kommentare-lesen "%DATEI%"
pause
goto menu

:oberflaeche
echo Starte die Oberflaeche in einem eigenen Fenster ...
start "Praesentationsassistent - Ueberarbeiten" "%PY%" -P -m praesentationsassistent.cli ueberarbeiten --ordner "%AUSGABE%"
timeout /t 2 /nobreak >nul
start "" http://127.0.0.1:8765
echo Browser sollte sich geoeffnet haben. Das andere Fenster (schwarz) muss
echo offen bleiben, solange du die Oberflaeche benutzt -- einfach schliessen,
echo wenn du fertig bist.
pause
goto menu

:ende
endlocal
exit /b 0
