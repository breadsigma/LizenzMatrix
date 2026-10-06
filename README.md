# LizenzMatrix

Überarbeitung der Lizenzmatrix: weniger Microsoft 365 E5, mehr Entra ID P2.

| Datei | Inhalt |
|---|---|
| [`docs/Lizenzmatrix-v2.md`](docs/Lizenzmatrix-v2.md) | Neue Matrix (Diagramm), Regeln, Vorgehen bei der Umstellung, PowerShell-Befehle |
| [`docs/Lizenzmatrix-v2.png`](docs/Lizenzmatrix-v2.png) | Diagramm als Bild |
| [`scripts/Get-LizenzReport.ps1`](scripts/Get-LizenzReport.ps1) | Liest alle Lizenzen und Nutzungsdaten aus und markiert E5-Kandidaten (nur lesend) |

## Report ausführen

Das Script meldet sich über **Az.Accounts** (Azure PowerShell) an und liest Microsoft Graph direkt per REST.
Das Microsoft.Graph-Modul wird nicht benötigt.

```powershell
Install-Module Az.Accounts -Scope CurrentUser      # einmalig
Install-Module ImportExcel -Scope CurrentUser      # optional, für eine .xlsx-Datei

cd C:\Pfad\zum\Script
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Get-LizenzReport.ps1                             # Standard: 90 Tage, Browser-Anmeldung
.\Get-LizenzReport.ps1 -IncludeTeamsPhone          # prüft zusätzlich Teams-Telefonie (Modul MicrosoftTeams)
```

Benötigte Rollen: **Global Reader** und **Reports Reader**.

Hinweise:
- Die normale Browser-Anmeldung verwenden. `-UseDeviceCode` nur, wenn kein Browserfenster erscheint;
  der Gerätecode-Login ist in vielen Organisationen per Conditional Access gesperrt
  („Hierauf haben Sie keinen Zugriff … Authentifizierungsflow“).
- Vorher im Microsoft 365 Admin Center unter *Einstellungen → Organisationseinstellungen → Berichte*
  die Option „Anonymisierte Benutzer-, Gruppen- und Websitenamen anzeigen“ **deaktivieren**.
  Sonst können die Nutzungsdaten den Benutzern nicht zugeordnet werden.
