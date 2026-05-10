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