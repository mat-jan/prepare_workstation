# ============================================================
#  CORPORATE PC PREPARATION SCRIPT / SKRYPT PRZYGOTOWANIA PC
# ============================================================

#region --- JĘZYK / LANGUAGE ---
Clear-Host
Write-Host "Select language / Wybierz język:" -ForegroundColor White
Write-Host "1. English"
Write-Host "2. Polski"
$choice = Read-Host "Choice / Wybór (1/2)"

$T = @{} # Obiekt tłumaczeń

if ($choice -eq "1") {
    $T.Step0 = "Step 0: Setting computer name"
    $T.Step1 = "Step 1: Power settings - Never sleep"
    $T.Step2 = "Step 2: Removing old Office versions"
    $T.Step3 = "Step 3: Application installation"
    $T.Step4 = "Step 4: Restoring ExecutionPolicy"
    $T.AdminReq = "Please run this script as Administrator!"
    $T.FolderErr = "Installer folder not found at: "
    $T.CompNamePrompt = "Current name: {0}`n`nEnter new computer name (max 15 chars, alphanumeric/hyphens):"
    $T.CompNameTitle = "Computer Name"
    $T.CompNameChanged = "Computer name changed to: {0} (takes effect after reboot)"
    $T.CompNameInvalid = "Invalid name! Skipping name change."
    $T.CompNameSkip = "Computer name unchanged: {0}"
    $T.PowerSS = "Screen saver: disabled"
    $T.PowerPlan = "Power plan: High Performance activated"
    $T.PowerPlanErr = "High Performance plan not found, editing active plan"
    $T.PowerMonitor = "Monitor timeout: never"
    $T.PowerStandby = "Standby: never"
    $T.PowerHib = "Hibernation: disabled"
    $T.OfficeSearch = "Searching for old Office installations..."
    $T.OfficeFound = "Found Office installations to remove:"
    $T.OfficeUninstalling = "Uninstalling: {0}"
    $T.OfficeDone = "Uninstalled: {0} (Code: {1})"
    $T.OfficeC2R = "Removing Office Click-to-Run..."
    $T.OfficeFolder = "Removed folder: {0}"
    $T.OfficeSkip = "No old Office found - skipping"
    $T.AppInstall = "Installing: {0}"
    $T.AppNotFound = "{0} - file not found: {1}"
    $T.AppManual = "-> Opening {0} installer (manual setup)..."
    $T.AppDone = "{0} installed (Code: {1})"
    $T.AppClosed = "{0} - installer closed"
    $T.ExecRestored = "ExecutionPolicy restored to: Restricted"
    $T.Summary = "DONE! Computer preparation finished."
    $T.RebootReq = "WARNING: Reboot required for computer name '{0}' to take effect!"
} else {
    $T.Step0 = "Krok 0: Ustawianie nazwy komputera"
    $T.Step1 = "Krok 1: Wygaszacz ekranu i hibernacja - nigdy"
    $T.Step2 = "Krok 2: Usuniecie starego Office"
    $T.Step3 = "Krok 3: Instalacja aplikacji"
    $T.Step4 = "Krok 4: Przywrocenie Execution Policy"
    $T.AdminReq = "Uruchom skrypt jako Administrator!"
    $T.FolderErr = "Folder z instalatorami nie istnieje: "
    $T.CompNamePrompt = "Aktualna nazwa: {0}`n`nWpisz nowa nazwe (max 15 znakow, litery/cyfry/myslniki):"
    $T.CompNameTitle = "Nazwa komputera"
    $T.CompNameChanged = "Nazwa komputera zmieniona na: {0} (wejdzie w zycie po restarcie)"
    $T.CompNameInvalid = "Nieprawidlowa nazwa! Pomijam zmiane."
    $T.CompNameSkip = "Nazwa komputera bez zmian: {0}"
    $T.PowerSS = "Wygaszacz ekranu: wylaczony"
    $T.PowerPlan = "Plan zasilania: Wysoka wydajnosc aktywny"
    $T.PowerPlanErr = "Nie znaleziono planu Wysoka wydajnosc, edytuje aktywny plan"
    $T.PowerMonitor = "Wygaszenie ekranu: nigdy"
    $T.PowerStandby = "Usypianie: nigdy"
    $T.PowerHib = "Hibernacja: wylaczona"
    $T.OfficeSearch = "Szukanie starych instalacji Microsoft Office..."
    $T.OfficeFound = "Znaleziono instalacje Office do usuniecia:"
    $T.OfficeUninstalling = "Odinstalowuje: {0}"
    $T.OfficeDone = "Odinstalowano: {0} (Kod: {1})"
    $T.OfficeC2R = "Usuwanie Office Click-to-Run..."
    $T.OfficeFolder = "Usunieto folder: {0}"
    $T.OfficeSkip = "Nie znaleziono starych instalacji Office - pomijam"
    $T.AppInstall = "Instalacja: {0}"
    $T.AppNotFound = "{0} - plik nie znaleziony: {1}"
    $T.AppManual = "-> Otwieranie instalatora {0} (instalacja reczna)..."
    $T.AppDone = "{0} zainstalowany (Kod: {1})"
    $T.AppClosed = "{0} - instalator zamkniety"
    $T.ExecRestored = "ExecutionPolicy przywrocona do: Restricted"
    $T.Summary = "GOTOWE! Przygotowanie komputera ukonczone."
    $T.RebootReq = "UWAGA: Wymagany restart aby nazwa '{0}' weszla w zycie!"
}
#endregion

#region --- KONFIGURACJA ---
$InstallerPath = "C:\instalki"
#endregion

function Write-Step { param($msg) Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Skip { param($msg) Write-Host "    [--] $msg" -ForegroundColor Yellow }
function Write-Err  { param($msg) Write-Host "    [!!] $msg" -ForegroundColor Red }

# Sprawdz uprawnienia
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Err $T.AdminReq
    pause; exit 1
}

# Sprawdz folder
if (-not (Test-Path $InstallerPath)) {
    Write-Err ($T.FolderErr + $InstallerPath)
    pause; exit 1
}

# --- KROK 0: NAZWA KOMPUTERA ---
Write-Step $T.Step0
Add-Type -AssemblyName Microsoft.VisualBasic
$currentName = $env:COMPUTERNAME
$newName = [Microsoft.VisualBasic.Interaction]::InputBox(($T.CompNamePrompt -f $currentName), $T.CompNameTitle, $currentName)

if ($newName -ne "" -and $newName -ne $currentName) {
    if ($newName -match "^[a-zA-Z0-9\-]{1,15}$") {
        Rename-Computer -NewName $newName -Force
        Write-OK ($T.CompNameChanged -f $newName)
    } else {
        Write-Err $T.CompNameInvalid
    }
} else {
    Write-Skip ($T.CompNameSkip -f $currentName)
}

# --- KROK 1: ZASILANIE ---
Write-Step $T.Step1
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveActive" -Value "0" -Force
Write-OK $T.PowerSS

$highPerfGUID = (powercfg /list | Select-String "Wysoka|High" | ForEach-Object { ($_ -split "\s+")[3] })
if ($highPerfGUID) {
    powercfg /setactive $highPerfGUID
    Write-OK $T.PowerPlan
} else {
    Write-Skip $T.PowerPlanErr
}

powercfg /change monitor-timeout-ac 0
powercfg /change monitor-timeout-dc 0
Write-OK $T.PowerMonitor

powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
Write-OK $T.PowerStandby

powercfg /hibernate off
Write-OK $T.PowerHib

# --- KROK 2: USUWANIE OFFICE ---
Write-Step $T.Step2
Write-Host "    $($T.OfficeSearch)"
$officeUninstallKeys = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
$officeFound = @()
foreach ($regPath in $officeUninstallKeys) {
    if (Test-Path $regPath) {
        Get-ChildItem $regPath | ForEach-Object {
            $displayName = ($_ | Get-ItemProperty -Name DisplayName -ErrorAction SilentlyContinue).DisplayName
            $uninstallStr = ($_ | Get-ItemProperty -Name UninstallString -ErrorAction SilentlyContinue).UninstallString
            if ($displayName -match "Microsoft 365|Microsoft Office|Office 1[456]" -and $uninstallStr) {
                $officeFound += [PSCustomObject]@{ Name = $displayName; Uninstall = $uninstallStr }
            }
        }
    }
}

if ($officeFound.Count -eq 0) {
    Write-Skip $T.OfficeSkip
} else {
    Write-Host "    $($T.OfficeFound)" -ForegroundColor Yellow
    foreach ($app in $officeFound) {
        Write-Host "    - $($app.Name)" -ForegroundColor Yellow
        Write-Host "    $($T.OfficeUninstalling -f $app.Name)" -ForegroundColor Magenta
        if ($app.Uninstall -match "^MsiExec") {
            $guid = ($app.Uninstall -replace "MsiExec.exe\s*/[IX]", "").Trim()
            $proc = Start-Process "msiexec.exe" -ArgumentList "/qn /x $guid /norestart" -Wait -PassThru
        } else {
            $proc = Start-Process "cmd.exe" -ArgumentList "/c $($app.Uninstall) /quiet /norestart" -Wait -PassThru
        }
        Write-OK ($T.OfficeDone -f $app.Name, $proc.ExitCode)
    }
}

# --- KROK 3: INSTALACJA APLIKACJI ---
function Install-App {
    param([string]$Name, [string]$File, [string]$SilentArgs = "", [bool]$Silent = $true)
    $fullPath = Join-Path $InstallerPath $File
    if (-not (Test-Path $fullPath)) { Write-Skip ($T.AppNotFound -f $Name, $File); return }
    Write-Step ($T.AppInstall -f $Name)
    if ($Silent -and $SilentArgs -ne "") {
        $proc = Start-Process -FilePath $fullPath -ArgumentList $SilentArgs -Wait -PassThru
        Write-OK ($T.AppDone -f $Name, $proc.ExitCode)
    } else {
        Write-Host "    $($T.AppManual -f $Name)" -ForegroundColor Magenta
        $proc = Start-Process -FilePath $fullPath -Wait -PassThru
        Write-OK ($T.AppClosed -f $Name)
    }
}

Install-App "AnyDesk" "AnyDesk.exe" "--install `"C:\Program Files (x86)\AnyDesk`" --silent --create-shortcuts --start-with-win" $true

$esetMsi = Join-Path $InstallerPath "ees_nt64.msi"
if (Test-Path $esetMsi) {
    Write-Step ($T.AppInstall -f "ESET")
    $proc = Start-Process "msiexec.exe" -ArgumentList "/qn /i `"$esetMsi`" ADDLOCAL=ALL REBOOT_WHEN_NEEDED=0" -Wait -PassThru
    Write-OK ($T.AppDone -f "ESET", $proc.ExitCode)
}

Install-App "Google Chrome" "googlechromestandaloneenterprise64.msi" "/qn /i" $true
Install-App "Intel DSA" "Intel-Driver-and-Support-Assistant-Installer.exe" "-q" $true
Install-App "Microsoft Office 32-bit PL" "OfficeSetup32bitPL.exe" "" $false

# --- KROK 4: PORZĄDKI ---
Write-Step $T.Step4
Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope LocalMachine -Force
Write-OK $T.ExecRestored

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  $($T.Summary)" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
if ($newName -ne "" -and $newName -ne $currentName) {
    Write-Host "`n  $($T.RebootReq -f $newName)" -ForegroundColor Yellow
}
pause