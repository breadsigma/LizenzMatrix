# LizenzMatrix

Überarbeitung der Lizenzmatrix des Kantons Zug: weniger Microsoft 365 E5, mehr Entra ID P2,
feste Regeln für den Entzug ungenutzter Lizenzen.

| Datei | Inhalt |
|---|---|
| [`docs/Lizenzmatrix-v2.md`](docs/Lizenzmatrix-v2.md) | Neue Matrix (Entscheidungsbaum), Voraussetzungen pro Persona, Betriebsregeln, Vorgehen bei der Umstellung, Gruppenstruktur |
| [`docs/Lizenzmatrix-v2.png`](docs/Lizenzmatrix-v2.png) | Entscheidungsbaum als Bild (Mermaid-Quelle im Markdown, in draw.io importierbar) |
| [`scripts/Get-LizenzReport.ps1`](scripts/Get-LizenzReport.ps1) | Liest Lizenzen, letzte Anmeldung und Nutzung aus Entra ID / Microsoft 365 und markiert Entzugs- und P2-Kandidaten. Nur lesend. |
| [`scripts/New-LizenzAuswertung.ps1`](scripts/New-LizenzAuswertung.ps1) | Wertet einen Report-Ordner aus: Kandidatenlisten (gesamt, pro Amt, pro EmployeeType), Sparpotenzial, HTML-Bericht mit Grafiken. |
| [`scripts/Check-LicenseGroupAD.ps1`](scripts/Check-LicenseGroupAD.ps1) | Einfache Variante für die ISE: prüft die Mitglieder der lokalen AD-Lizenzgruppe mit `Get-ADUser` (Domänen-Anmeldung), ohne Cloud. |

Die Auswertungen mit Personendaten (Berichte, Kandidatenlisten) werden **nicht** im Repo abgelegt.

## Report ausführen

Das Script braucht **keine Zusatzmodule**. Es öffnet den Browser, meldet sich wie Azure CLI /
Azure PowerShell an (Microsoft-eigene Apps, keine Admin-Freigabe nötig) und liest Microsoft Graph
direkt per REST. Läuft in Windows PowerShell 5.1 und PowerShell 7.

```powershell
cd C:\Pfad\zum\Script
Unblock-File .\Get-LizenzReport.ps1
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Standard: Browser öffnet sich auf diesem PC
.\Get-LizenzReport.ps1 -Account admin@zg.ch -Client AzureCli

# Auf einem Server (Remotedesktop): Link auf dem eigenen PC öffnen, Antwort-Adresse zurück einfügen
.\Get-LizenzReport.ps1 -Account admin@zg.ch -Client AzureCli -NoBrowser

# Mit Nutzungsberichten (braucht Admin Consent für "Microsoft Graph Command Line Tools")
.\Get-LizenzReport.ps1 -Account admin@zg.ch -Client GraphCli -NoBrowser

# Report und Auswertung in einem Schritt
.\Get-LizenzReport.ps1 -Account admin@zg.ch -Client AzureCli -NoBrowser -Auswertung

# Auswertung eines bestehenden Reports, mit Vertragspreis
.\New-LizenzAuswertung.ps1 -ReportPath .\LizenzReport_2026-10-06 -PreisE5 38.20
```

## Variante ohne Cloud: AD-Lizenzgruppe prüfen

`Check-LicenseGroupAD.ps1` läuft in der PowerShell ISE auf einem Server mit AD-Modul (RSAT). Oben im Script
den Gruppennamen eintragen, F5. Es liest pro Mitglied `Enabled`, `whenCreated`, `LastLogonDate`,
`AccountExpirationDate` und ordnet Status und Massnahme zu. Ausgabe nach `C:\Temp\Lizenzcheck_<Datum>`.
Der Entzug ist als `-WhatIf`-Block auskommentiert.

Hinweis: `LastLogonDate` ist die letzte **Domänen**-Anmeldung (Laptop, VDI), nicht die Cloud-Anmeldung.
Für Konten ohne Kantonslaptop ist der Cloud-Report die verlässlichere Quelle.

### Benötigte Rollen (vorher in PIM aktivieren)

| Daten | Rolle | Anmelde-App |
|---|---|---|
| Lizenzen, Benutzer, Gruppenzuweisung | User Administrator oder Global Reader | `AzureCli` |
| Letzte Anmeldung (`signInActivity`) | **Security Administrator**, Security Reader oder Global Reader | `AzureCli` |
| Nutzungsberichte (Mail, Teams, OneDrive, Office-Aktivierungen) | **Reports Reader** oder Global Reader | `GraphCli` (einmalig Admin Consent für `Reports.Read.All`, `ReportSettings.Read.All`) |

### Ausgabe

Ordner `LizenzReport_<Datum>` mit:

- `01_SKU-Uebersicht.csv`: gekaufte, zugewiesene und freie Lizenzen pro Produkt
- `02_Benutzer-Lizenzen.csv`: ein Eintrag pro Benutzer mit Lizenzen, Zuweisung (Gruppe/direkt), letzter Anmeldung,
  Nutzung, Spalte **Anmeldestatus** und Spalte **Empfehlung**
- `03_E5-Kandidaten.csv`: nur die E5-Benutzer mit einer Massnahme
- `Rohdaten\`: die originalen Nutzungsberichte

Werte in der Spalte *Empfehlung*:

| Empfehlung | Bedeutung |
|---|---|
| `Lizenz entfernen` | Konto deaktiviert (Austritt) oder über 365 Tage ohne Anmeldung |
| `Lizenz entfernen / pruefen` | 91–365 Tage ohne Anmeldung, oder nie angemeldet bei Konto älter als 90 Tage |
| `Kein Entzug (vorbereiteter Eintritt)` | Konto deaktiviert, jünger als 60 Tage, nie angemeldet |
| `Abwarten (neues Konto)` | Konto aktiv, jünger als 90 Tage, noch nie angemeldet |
| `Kandidat E5 -> Entra ID P2` | Aktiv, aber keine Nutzung von Mail/Teams/OneDrive/Office (nur mit Nutzungsberichten) |
| `Pruefen: E5 -> kleinere Lizenz` | Nur Web-Nutzung, keine Office-Desktop-Apps (nur mit Nutzungsberichten) |
| `E5 behalten` | Nutzung nachgewiesen |
| `Unbestimmt (keine Nutzungsdaten)` | Aktiv, Nutzungsberichte nicht verfügbar |

Parameter: `-InactiveDays 90`, `-NewAccountDays 60`, `-Period D90`, `-IncludeTeamsPhone`, `-UseDeviceCode`, `-TenantId`.

### Hinweise

- `-UseDeviceCode` ist in vielen Organisationen per Conditional Access gesperrt
  („Hierauf haben Sie keinen Zugriff … Authentifizierungsflow“). Standard ist die Browser-Anmeldung.
- Im Microsoft 365 Admin Center unter *Einstellungen → Organisationseinstellungen → Berichte* muss
  „Anonymisierte Benutzer-, Gruppen- und Websitenamen anzeigen“ **deaktiviert** sein, sonst lassen sich die
  Nutzungsberichte nicht zuordnen. Das Script prüft das und warnt.
- Optional `Install-Module ImportExcel -Scope CurrentUser` für eine zusätzliche `.xlsx`-Datei.
