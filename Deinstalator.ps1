# ============================================================
#  UNINSTALLER - TEST MACHINE / DEINSTALATOR - MASZYNA TESTOWA
# ============================================================

#region --- JĘZYK / LANGUAGE ---
Clear-Host
Write-Host "Select language / Wybierz język:" -ForegroundColor White
Write-Host "1. English"
Write-Host "2. Polski"
$choice = Read-Host "Choice / Wybór (1/2)"

$T = @{} # Dictionary for translations

if ($choice -eq "1") {
    $T.AdminReq = "Please run this script as Administrator!"
    $T.BannerHeader = "UNINSTALLER - TEST MACHINE"
    $T.BannerDesc = "Removes: AnyDesk, ESET, Chrome, Intel DSA, Office`nand reverts system settings."
    $T.StepUninstall = "Uninstalling: {0}"
    $T.Found = "Found: {0}"
    $T.Removed = "{0} removed (Code: {1})"
    $T.NotFound = "{0} - not found in the system"
    $T.FolderCleanup = "Removed folder: {0}"
    $T.OfficeC2R = "Removing Office Click-to-Run..."
    $T.PowerStep = "Reverting power settings to default"
    $T.PowerBalanced = "Power plan: Balanced restored"
    $T.PowerMonitor = "Monitor timeout: 10m (AC), 5m (DC)"
    $T.PowerStandby = "Sleep: 30m (AC), 15m (DC)"
    $T.PowerSS = "Screensaver: restored (10 min)"
    $T.PowerHib = "Hibernation: restored"
    $T.ExecStep = "Restoring ExecutionPolicy to Restricted"
    $T.ExecRestored = "ExecutionPolicy: Restricted"
    $T.SummaryHeader = "DONE! Test machine cleaned."
    $T.SummaryBody = "You can now run Setup-Corporate.ps1 to start fresh."
} else {
    $T.AdminReq = "Uruchom skrypt jako Administrator!"
    $T.BannerHeader = "DEINSTALATOR - MASZYNA TESTOWA"
    $T.BannerDesc = "Usuwa: AnyDesk, ESET, Chrome, Intel DSA, Office`noraz cofa ustawienia systemu."
    $T.StepUninstall = "Odinstalowywanie: {0}"
    $T.Found = "Znaleziono: {0}"
    $T.Removed = "{0} usunięto (Kod: {1})"
    $T.NotFound = "{0} - nie znaleziono w systemie"
    $T.FolderCleanup = "Usunięto folder: {0}"
    $T.OfficeC2R = "Usuwanie Office Click-to-Run..."
    $T.PowerStep = "Cofanie ustawień zasilania do domyślnych"
    $T.PowerBalanced = "Plan zasilania: Zrównoważony przywrócony"
    $T.PowerMonitor = "Wygaszenie ekranu: 10m (AC), 5m (DC)"
    $T.PowerStandby = "Usypianie: 30m (AC), 15m (DC)"
    $T.PowerSS = "Wygaszacz ekranu: przywrócony (10 min)"
    $T.PowerHib = "Hibernacja: przywrócona"
    $T.ExecStep = "Przywracanie ExecutionPolicy do Restricted"
    $T.ExecRestored = "ExecutionPolicy: Restricted"
    $T.SummaryHeader = "GOTOWE! Maszyna testowa wyczyszczona."
    $T.SummaryBody = "Możesz uruchomić Setup-Firmowy.ps1, aby zacząć od nowa."
}
#endregion

# Kolory w konsoli
function Write-Step { param($msg) Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Skip { param($msg) Write-Host "    [--] $msg" -ForegroundColor Yellow }
function Write-Err  { param($msg) Write-Host "    [!!] $msg" -ForegroundColor Red }

# Sprawdź uprawnienia
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Err $T.AdminReq
    pause; exit 1
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Red
Write-Host "  $($T.BannerHeader)" -ForegroundColor Red
Write-Host "  $($T.BannerDesc)" -ForegroundColor Red
Write-Host "============================================" -ForegroundColor Red

# --- FUNKCJA POMOCNICZA ---
function Uninstall-ByName {
    param([string]$Pattern, [string]$DisplayLabel)
    Write-Step ($T.StepUninstall -f $DisplayLabel)
    $regPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
    $found = $false
    foreach ($regPath in $regPaths) {
        if (-not (Test-Path $regPath)) { continue }
        Get-ChildItem $regPath | ForEach-Object {
            $props = $_ | Get-ItemProperty -ErrorAction SilentlyContinue
            if ($props.DisplayName -match $Pattern -and $props.UninstallString) {
                $found = $true
                Write-Host "    $($T.Found -f $props.DisplayName)" -ForegroundColor Magenta
                if ($props.UninstallString -match "MsiExec") {
                    $guid = ($props.UninstallString -replace "MsiExec.exe\s*/[IX]", "").Trim()
                    $proc = Start-Process "msiexec.exe" -ArgumentList "/qn /x $guid /norestart" -Wait -PassThru
                } else {
                    $proc = Start-Process "cmd.exe" -ArgumentList "/c $($props.UninstallString) /S /silent /quiet /norestart" -Wait -PassThru
                }
                Write-OK ($T.Removed -f $props.DisplayName, $proc.ExitCode)
            }
        }
    }
    if (-not $found) { Write-Skip ($T.NotFound -f $DisplayLabel) }
}

# 1. ANYDESK
Write-Step ($T.StepUninstall -f "AnyDesk")
$anydesk = "C:\Program Files (x86)\AnyDesk\AnyDesk.exe"
if (Test-Path $anydesk) {
    $proc = Start-Process $anydesk -ArgumentList "--remove" -Wait -PassThru
    Write-OK ($T.Removed -f "AnyDesk", $proc.ExitCode)
} else {
    Uninstall-ByName -Pattern "AnyDesk" -DisplayLabel "AnyDesk (Registry)"
}

# 2. ESET
Uninstall-ByName -Pattern "ESET" -DisplayLabel "ESET"

# 3. CHROME
Uninstall-ByName -Pattern "Google Chrome" -DisplayLabel "Google Chrome"
$chromePaths = @("$env:ProgramFiles\Google\Chrome", "${env:ProgramFiles(x86)}\Google\Chrome")
foreach ($p in $chromePaths) {
    if (Test-Path $p) { Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue; Write-OK ($T.FolderCleanup -f $p) }
}

# 4. INTEL DSA
Uninstall-ByName -Pattern "Intel.*(Driver|DSA|Support Assistant)" -DisplayLabel "Intel Driver and Support Assistant"

# 5. OFFICE
Write-Step ($T.StepUninstall -f "Microsoft Office")
$officeFound = @()
$officeKeys = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
foreach ($regPath in $officeKeys) {
    if (Test-Path $regPath) {
        Get-ChildItem $regPath | ForEach-Object {
            $props = $_ | Get-ItemProperty -ErrorAction SilentlyContinue
            if ($props.DisplayName -match "Microsoft 365|Microsoft Office|Office 1[456]" -and $props.UninstallString) {
                $officeFound += [PSCustomObject]@{ Name = $props.DisplayName; Uninstall = $props.UninstallString }
            }
        }
    }
}

if ($officeFound.Count -eq 0) { Write-Skip ($T.NotFound -f "Office") }
else {
    foreach ($app in $officeFound) {
        Write-Host "    $($T.StepUninstall -f $app.Name)" -ForegroundColor Magenta
        if ($app.Uninstall -match "MsiExec") {
            $guid = ($app.Uninstall -replace "MsiExec.exe\s*/[IX]", "").Trim()
            $proc = Start-Process "msiexec.exe" -ArgumentList "/qn /x $guid /norestart" -Wait -PassThru
        } else {
            $proc = Start-Process "cmd.exe" -ArgumentList "/c $($app.Uninstall) /quiet /norestart" -Wait -PassThru
        }
        Write-OK ($T.Removed -f $app.Name, $proc.ExitCode)
    }
}

$c2rSetup = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"
if (Test-Path $c2rSetup) {
    Write-Host "    $($T.OfficeC2R)" -ForegroundColor Magenta
    Start-Process $c2rSetup -ArgumentList "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=O365ProPlusRetail.16_pl-pl_x-none" -Wait
}

# 6. POWER SETTINGS
Write-Step $T.PowerStep
$balancedGUID = (powercfg /list | Select-String "Zrownow|Balanced" | ForEach-Object { ($_ -split "\s+")[3] })
if ($balancedGUID) { powercfg /setactive $balancedGUID; Write-OK $T.PowerBalanced }

powercfg /change monitor-timeout-ac 10
powercfg /change monitor-timeout-dc 5
Write-OK $T.PowerMonitor

powercfg /change standby-timeout-ac 30
powercfg /change standby-timeout-dc 15
Write-OK $T.PowerStandby

Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveActive" -Value "1" -Force
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveTimeOut" -Value "600" -Force
Write-OK $T.PowerSS

powercfg /hibernate on
Write-OK $T.PowerHib

# 7. EXECUTION POLICY
Write-Step $T.ExecStep
Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope LocalMachine -Force
Write-OK $T.ExecRestored

# SUMMARY
Write-Host "`n============================================" -ForegroundColor Green
Write-Host "  $($T.SummaryHeader)" -ForegroundColor Green
Write-Host "  $($T.SummaryBody)" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
pause