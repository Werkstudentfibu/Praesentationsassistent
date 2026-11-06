<#
.SYNOPSIS
  Richtet den Praesentationsassistenten auf diesem Windows-Rechner ein:
  Python pruefen/installieren, venv anlegen, Programm installieren,
  Einstellungsdatei + Desktop-Verknuepfung anlegen.

.DESCRIPTION
  Einmalig ausfuehren, direkt aus dem kopierten Projektordner (der Ordner
  mit pyproject.toml, praesentationsassistent\, markenprofile\, windows\).
  KEIN GitHub-Zugang noetig -- der Code liegt schon lokal vor (Testphase,
  Upload/Selbst-Update folgt erst spaeter).

  Aufruf (PowerShell, im Ordner "windows" dieses Projekts):
    powershell -ExecutionPolicy Bypass -File .\einrichten.ps1

  Aufbau danach (Projektordner = ein Verzeichnis ueber "windows"):
    <Projektordner>\.venv\                  Python-Umgebung
    <Projektordner>\praesentation.env        eigene Einstellungen (Langdock-Key)
    <Projektordner>\windows\start.cmd        Programm starten
    Desktop-Verknuepfung "Praesentationsassistent"
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$Hier = Split-Path -Parent $MyInvocation.MyCommand.Path
$Root = Split-Path -Parent $Hier

$LogFile = Join-Path $Root "einrichten.log"
try { Start-Transcript -Path $LogFile -Force | Out-Null } catch { }

function Schritt([string]$t) { Write-Host ""; Write-Host "==> $t" -ForegroundColor Cyan }
function Ok([string]$t)      { Write-Host "    [ok] $t" -ForegroundColor Green }
function Hinweis([string]$t) { Write-Host "    [!]  $t" -ForegroundColor Yellow }
function Abbruch([string]$text, [string]$abhilfe = "") {
    Write-Host ""
    Write-Host "FEHLER: $text" -ForegroundColor Red
    if ($abhilfe) {
        Write-Host ""; Write-Host "Was tun?" -ForegroundColor Yellow
        foreach ($z in $abhilfe -split "`n") { Write-Host "  $z" }
    }
    Write-Host ""; Write-Host "Protokoll: $LogFile"
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}
function Aufruf([string]$exe, [string[]]$argumente) {
    $alt = $ErrorActionPreference; $ErrorActionPreference = "Continue"
    try { & $exe @argumente | Out-Host; return $LASTEXITCODE }
    catch { Write-Host "    $($_.Exception.Message)"; return 9999 }
    finally { $ErrorActionPreference = $alt }
}
function PfadNeuLaden {
    $m = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $u = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$m;$u"
}
function FindePython {
    # Liefert den vollen Pfad zu einem funktionierenden Python >= 3.11 oder $null.
    # Jeder Kandidat wird wirklich ausgefuehrt: "python" ist unter Windows oft
    # nur der Microsoft-Store-Platzhalter (WindowsApps), den Get-Command zwar
    # findet, der aber bei jedem Aufruf mit Code 9009 abbricht.
    $alt = $ErrorActionPreference; $ErrorActionPreference = "Continue"
    try {
        foreach ($kandidat in @("py -3", "python")) {
            $teile = $kandidat -split " "
            $exe = $teile[0]
            $vorArgs = @($teile | Select-Object -Skip 1)
            if (-not (Get-Command $exe -ErrorAction SilentlyContinue)) { continue }
            try {
                $ausgabe = & $exe @vorArgs -c "import sys; print(sys.executable if sys.version_info >= (3, 11) else '')" 2>$null
            } catch { continue }
            if ($LASTEXITCODE -ne 0 -or -not $ausgabe) { continue }
            $pfad = "$($ausgabe | Select-Object -Last 1)".Trim()
            if ($pfad -and (Test-Path $pfad)) { return $pfad }
        }
        return $null
    } finally { $ErrorActionPreference = $alt }
}

if (-not (Test-Path (Join-Path $Root "pyproject.toml"))) {
    Abbruch "pyproject.toml nicht gefunden in $Root." "Dieses Skript muss im Ordner 'windows' INNERHALB des kopierten Projektordners liegen."
}

# ---------------------------------------------------------------------------
# 1. Python
# ---------------------------------------------------------------------------
Schritt "Python pruefen"
$PyBefehl = FindePython
if (-not $PyBefehl) {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Abbruch "Python (ab Version 3.11) fehlt und winget ist nicht verfuegbar." "Python 3.12 von https://www.python.org/downloads/ installieren (Haken bei 'Add python.exe to PATH' setzen), danach dieses Skript erneut starten."
    }
    Write-Host "         Installiere Python per winget ..."
    $rc = Aufruf "winget" @("install", "-e", "--id", "Python.Python.3.12", "--accept-package-agreements", "--accept-source-agreements")
    if ($rc -ne 0) { Hinweis "winget meldete Code $rc" }
    PfadNeuLaden
    $PyBefehl = FindePython
    if (-not $PyBefehl) {
        Abbruch "Python ist nach der Installation nicht auffindbar." "Neues PowerShell-Fenster oeffnen und das Skript erneut starten."
    }
}
Ok ((& $PyBefehl --version) -join " ")

# ---------------------------------------------------------------------------
# 2. venv + Programm installieren
# ---------------------------------------------------------------------------
Schritt "Python-Umgebung anlegen (.venv)"
$Venv = Join-Path $Root ".venv"
if (-not (Test-Path (Join-Path $Venv "Scripts\python.exe"))) {
    $rc = Aufruf $PyBefehl @("-m", "venv", $Venv)
    if ($rc -ne 0) { Abbruch "venv konnte nicht angelegt werden (Code $rc)." }
    Ok "venv angelegt in $Venv"
} else {
    Ok "venv existiert bereits"
}
$Py = Join-Path $Venv "Scripts\python.exe"

Schritt "Programm installieren (pip install .)"
$rc = Aufruf $Py @("-m", "pip", "install", "--upgrade", "pip")
$rc = Aufruf $Py @("-m", "pip", "install", $Root)
if ($rc -ne 0) { Abbruch "Installation ist fehlgeschlagen (Code $rc)." "Internet pruefen -- pip laedt die beiden Bausteine python-pptx/pillow aus dem PyPI-Index." }
Ok "Programm installiert"

# ---------------------------------------------------------------------------
# 3. Einstellungen
# ---------------------------------------------------------------------------
Schritt "Einstellungsdatei anlegen"
$Env = Join-Path $Root "praesentation.env"
if (-not (Test-Path $Env)) {
    Copy-Item (Join-Path $Root "praesentation.env.example") $Env
    Ok "praesentation.env aus der Vorlage angelegt -- eigenen Langdock-Key dort eintragen"
} else {
    Ok "praesentation.env bleibt wie sie ist"
}

# ---------------------------------------------------------------------------
# 4. Starter + Desktop-Verknuepfung
# ---------------------------------------------------------------------------
Schritt "Verknuepfung anlegen"
$StartCmd = Join-Path $Hier "start.cmd"
if (-not (Test-Path $StartCmd)) {
    Abbruch "$StartCmd fehlt." "windows\start.cmd fehlt im kopierten Projektordner."
}
try {
    $sh = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut((Join-Path ([Environment]::GetFolderPath("Desktop")) "Praesentationsassistent.lnk"))
    $lnk.TargetPath = $StartCmd
    $lnk.WorkingDirectory = $Root
    $lnk.Save()
    Ok "Verknuepfung 'Praesentationsassistent' auf dem Desktop"
} catch {
    Hinweis "Desktop-Verknuepfung nicht angelegt: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Fertig. Vor dem ersten Start:" -ForegroundColor Green
Write-Host "  1. $Env oeffnen und den eigenen Langdock-API-Key eintragen (LANGDOCK_API_KEY=...)."
Write-Host "  2. Start: Desktop-Verknuepfung 'Praesentationsassistent' bzw. $StartCmd"
try { Stop-Transcript | Out-Null } catch { }
exit 0
