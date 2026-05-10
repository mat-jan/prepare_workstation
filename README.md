# Skrypty przygotowania komputera firmowego

Zestaw skryptow PowerShell do automatycznego przygotowania nowego komputera w srodowisku firmowym oraz do resetowania maszyny testowej.

---

## Zawartosc repozytorium

```
/
├── Setup-Firmowy.ps1       # Glowny skrypt instalacyjny
├── Deinstalator.ps1        # Skrypt do czyszczenia maszyny testowej
├── Uruchom-Setup.bat       # Uruchamia Setup-Firmowy.ps1 jako Administrator
├── Uruchom-Deinstalator.bat# Uruchamia Deinstalator.ps1 jako Administrator
└── README.md
```

---

## Wymagania

- Windows 10 / Windows 11
- Konto z uprawnieniami **Administratora**
- Folder `C:\instalki\` z przygotowanymi instalatorami (patrz nizej)

---

## Przygotowanie folderu z instalatorami

Przed uruchomieniem skryptu utworz folder `C:\instalki\` i wgraj do niego nastepujace pliki:

| Plik | Aplikacja |
|------|-----------|
| `AnyDesk.exe` | AnyDesk |
| `ees_nt64.msi` | ESET Endpoint Security |
| `googlechromestandaloneenterprise64.msi` | Google Chrome Enterprise |
| `Intel-Driver-and-Support-Assistant-Installer.exe` | Intel Driver & Support Assistant |
| `OfficeSetup32bitPL.exe` | Microsoft Office 32-bit PL |

Nazwy plikow musza byc dokladnie takie jak powyzej (skrypt szuka ich po nazwie).

---

## Uruchomienie

### Instalacja (nowy komputer firmowy)

1. Skopiuj wszystkie pliki ze skryptami do dowolnego folderu (np. razem z `C:\instalki\`)
2. Kliknij **prawym przyciskiem myszy** na `Uruchom-Setup.bat`
3. Wybierz **"Uruchom jako administrator"**

### Deinstalacja (reset maszyny testowej)

1. Kliknij **prawym przyciskiem myszy** na `Uruchom-Deinstalator.bat`
2. Wybierz **"Uruchom jako administrator"**

> Nie uruchamiaj plikow `.ps1` bezposrednio — pliki `.bat` automatycznie ustawiaja wymagane uprawnienia i ExecutionPolicy.

---

## Co robi Setup-Firmowy.ps1

### Krok 0 — Nazwa komputera
- Wyswietla okno dialogowe z aktualna nazwa komputera
- Pozwala wpisac nowa nazwe (max 15 znakow, tylko litery/cyfry/myslniki)
- Zmiana wchodzi w zycie po restarcie

### Krok 1 — Ustawienia zasilania
- Wylacza wygaszacz ekranu
- Aktywuje plan zasilania "Wysoka wydajnosc"
- Ustawia wygaszenie ekranu i usypianie na **nigdy** (AC i DC)
- Wylacza hibernacje

### Krok 2 — Usuniecie starego Office
- Automatycznie wykrywa i usuwa stare instalacje Microsoft Office (wersje 14/15/16, Microsoft 365)
- Obsluguje instalacje MSI oraz Click-to-Run
- Czyści pozostale foldery Office

### Krok 3 — Instalacja aplikacji

| Aplikacja | Tryb |
|-----------|------|
| AnyDesk | Silent (bez okienek) |
| ESET Endpoint Security | Silent (bez okienek) |
| Google Chrome Enterprise | Silent (bez okienek) |
| Intel Driver & Support Assistant | Silent (bez okienek) |
| Microsoft Office 32-bit PL | Reczna (otwiera instalator) |

### Krok 4 — Przywrocenie ExecutionPolicy
- Po zakonczeniu instalacji przywraca `ExecutionPolicy` do `Restricted`

---

## Co robi Deinstalator.ps1

Usuwa wszystkie aplikacje zainstalowane przez `Setup-Firmowy.ps1` i cofa ustawienia systemowe:

- Odinstalowuje: AnyDesk, ESET, Google Chrome, Intel DSA, Microsoft Office
- Cofa ustawienia zasilania do domyslnych Windows (ekran: 10 min, uśpienie: 30 min, hibernacja wlaczona)
- Przywraca ExecutionPolicy do `Restricted`

---

## Dodawanie kolejnych aplikacji

W pliku `Setup-Firmowy.ps1` znajdz sekcje z komentarzem:

```
>> DODAJ KOLEJNE APLIKACJE TUTAJ <<
```

Przykladowe uzycie:

```powershell
# Silent .exe (np. 7-Zip)
Install-App -Name "7-Zip" -File "7z2301-x64.exe" -SilentArgs "/S" -Silent $true

# Silent .msi
$msi = Join-Path $InstallerPath "program.msi"
Start-Process "msiexec.exe" -ArgumentList "/qn /i `"$msi`" /norestart" -Wait

# Reczna instalacja
Install-App -Name "Program" -File "setup.exe" -Silent $false
```

---

## Aktywacja licencji ESET

W sekcji ESET w skrypcie znajdz linie:

```powershell
-ArgumentList "/qn /i `"$esetMsi`" ADDLOCAL=ALL REBOOT_WHEN_NEEDED=0"
```

I dodaj na koncu parametr z kluczem:

```powershell
-ArgumentList "/qn /i `"$esetMsi`" ADDLOCAL=ALL REBOOT_WHEN_NEEDED=0 ACTIVATION_DATA=key:AAAA-BBBB-CCCC-DDDD-EEEE"
```

---

## Najczestsze problemy

**Skrypt nie uruchamia sie**
Upewnij sie ze uruchamiasz przez `.bat`, a nie bezposrednio plik `.ps1`. Jezeli mimo to nie dziala, uruchom PowerShell jako Administrator i wpisz:
```powershell
Set-ExecutionPolicy RemoteSigned -Scope LocalMachine
```

**Nie znaleziono pliku instalatora**
Sprawdz czy nazwy plikow w folderze `C:\instalki\` sa identyczne jak w tabeli powyzej (wielkosc liter ma znaczenie).

**ESET zwraca blad instalacji**
Upewnij sie ze poprzednia wersja ESET zostala calkowicie odinstalowana przed uruchomieniem skryptu.

**Office nie instaluje sie po usunieciu starego**
Zrestartuj komputer po kroku deinstalacji i uruchom skrypt ponownie.

---

## Licencja

Skrypty sa dostepne do dowolnego uzytku w srodowiskach firmowych.



Here is the English version of your documentation. I’ve refined the phrasing to sound professional yet accessible, ensuring it’s clear for any IT administrator.

---

# Corporate PC Provisioning Scripts

A collection of PowerShell scripts designed for the automated setup of new corporate workstations and for resetting test machines to a clean state.

---

## Repository Structure

```
/
├── Setup-Corporate.ps1      # Main installation script
├── Uninstaller.ps1          # Script for cleaning up test machines
├── Run-Setup.bat            # Launches Setup-Corporate.ps1 as Administrator
├── Run-Uninstaller.bat      # Launches Uninstaller.ps1 as Administrator
└── README.md                # Documentation

```

---

## Requirements

* **Windows 10 / Windows 11**
* **Administrator** account privileges
* A folder named `C:\installers\` containing the required installation files (see below)

---

## Preparing the Installation Folder

Before running the script, create the folder `C:\installers\` and place the following files inside:

| File Name | Application |
| --- | --- |
| `AnyDesk.exe` | AnyDesk |
| `ees_nt64.msi` | ESET Endpoint Security |
| `googlechromestandaloneenterprise64.msi` | Google Chrome Enterprise |
| `Intel-Driver-and-Support-Assistant-Installer.exe` | Intel Driver & Support Assistant |
| `OfficeSetup32bitPL.exe` | Microsoft Office 32-bit (Polish) |

> **Note:** File names must match the table exactly, as the script identifies them by name.

---

## Usage

### Installation (New Corporate PC)

1. Copy all script files to any folder (e.g., alongside `C:\installers\`).
2. **Right-click** on `Run-Setup.bat`.
3. Select **"Run as administrator"**.

### Uninstallation (Test Machine Reset)

1. **Right-click** on `Run-Uninstaller.bat`.
2. Select **"Run as administrator"**.

> **Pro Tip:** Do not run the `.ps1` files directly. The `.bat` files automatically handle the necessary permissions and the `ExecutionPolicy` bypass for you.

---

## What Setup-Corporate.ps1 Does

### Step 0 — Computer Naming

* Displays a dialog box showing the current computer name.
* Allows you to enter a new name (max 15 characters, alphanumeric and hyphens only).
* The name change takes effect after a reboot.

### Step 1 — Power Settings

* Disables the screensaver.
* Activates the **"High Performance"** power plan.
* Sets screen timeout and sleep mode to **"Never"** (for both AC and Battery).
* Disables Hibernation.

### Step 2 — Legacy Office Removal

* Automatically detects and removes old Microsoft Office installations (versions 14/15/16, and Microsoft 365).
* Supports both MSI and Click-to-Run installations.
* Cleans up residual Office folders.

### Step 3 — Application Installation

| Application | Mode |
| --- | --- |
| AnyDesk | Silent (Background) |
| ESET Endpoint Security | Silent (Background) |
| Google Chrome Enterprise | Silent (Background) |
| Intel Driver & Support Assistant | Silent (Background) |
| Microsoft Office 32-bit | Manual (Opens the installer UI) |

### Step 4 — Security Cleanup

* Restores the system `ExecutionPolicy` to `Restricted` once the process is complete.

---

## What Uninstaller.ps1 Does

This script reverts the changes made by the setup script:

* **Uninstalls:** AnyDesk, ESET, Google Chrome, Intel DSA, and Microsoft Office.
* **Resets Power Settings:** Returns to Windows defaults (Screen: 10 min, Sleep: 30 min, Hibernation enabled).
* **Restores Security:** Resets `ExecutionPolicy` to `Restricted`.

---

## Adding New Applications

In `Setup-Corporate.ps1`, locate the section marked:

```powershell
# >> ADD ADDITIONAL APPLICATIONS HERE <<

```

### Examples:

**Silent .exe (e.g., 7-Zip)**

```powershell
Install-App -Name "7-Zip" -File "7z2301-x64.exe" -SilentArgs "/S" -Silent $true

```

**Silent .msi**

```powershell
$msi = Join-Path $InstallerPath "program.msi"
Start-Process "msiexec.exe" -ArgumentList "/qn /i `"$msi`" /norestart" -Wait

```

**Manual Installation**

```powershell
Install-App -Name "Custom Program" -File "setup.exe" -Silent $false

```

---

## ESET License Activation

To automate ESET activation, find the ESET section in the script and update the arguments:

**Find:**

```powershell
-ArgumentList "/qn /i `"$esetMsi`" ADDLOCAL=ALL REBOOT_WHEN_NEEDED=0"

```

**Change to:**

```powershell
-ArgumentList "/qn /i `"$esetMsi`" ADDLOCAL=ALL REBOOT_WHEN_NEEDED=0 ACTIVATION_DATA=key:AAAA-BBBB-CCCC-DDDD-EEEE"

```

---

## Troubleshooting

* **Script won't run:** Ensure you are using the `.bat` file. If issues persist, open PowerShell as Admin and run: `Set-ExecutionPolicy RemoteSigned -Scope LocalMachine`.
* **Installer not found:** Double-check that files in `C:\installers\` match the names in the script exactly.
* **ESET installation error:** Ensure any previous antivirus software (including older ESET versions) is fully removed before running the script.
* **Office installation fails:** If you just removed an old version of Office, a reboot is often required before the new one can be installed.

---

## License

These scripts are provided for free use within corporate environments. Use responsibly!