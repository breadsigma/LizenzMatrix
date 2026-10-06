# LizenzMatrix

Überarbeitung der Lizenzmatrix: weniger Microsoft 365 E5, mehr Entra ID P2.

| Datei | Inhalt |
|---|---|
| [`docs/Lizenzmatrix-v2.md`](docs/Lizenzmatrix-v2.md) | Neue Matrix (Diagramm), Regeln, Vorgehen bei der Umstellung, PowerShell-Befehle |
| [`docs/Lizenzmatrix-v2.png`](docs/Lizenzmatrix-v2.png) | Diagramm als Bild |
| [`scripts/Get-LizenzReport.ps1`](scripts/Get-LizenzReport.ps1) | Liest alle Lizenzen und Nutzungsdaten aus und markiert E5-Kandidaten (nur lesend) |

## Report ausführen

Das Script braucht **keine Zusatzmodule** (weder Microsoft.Graph noch Az). Es öffnet den Browser,
meldet sich wie Azure PowerShell an (Microsoft-eigene App, keine Admin-Freigabe nötig) und liest
Microsoft Graph direkt per REST. Läuft in Windows PowerShell 5.1 und PowerShell 7.

```powershell
cd C:\Pfad\zum\Script
Unblock-File .\Get-LizenzReport.ps1
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Get-LizenzReport.ps1                             # Standard: 90 Tage, Browser-Anmeldung
.\Get-LizenzReport.ps1 -RequestScopes              # falls Berechtigungen im Token fehlen (Zustimmungsdialog)
.\Get-LizenzReport.ps1 -IncludeTeamsPhone          # prüft zusätzlich Teams-Telefonie (Modul MicrosoftTeams)
```

Benötigte Rollen: **Global Reader** und **Reports Reader**.

Hinweise:
- `-UseDeviceCode` nur, wenn kein Browserfenster erscheint; der Gerätecode-Login ist in vielen
  Organisationen per Conditional Access gesperrt („Hierauf haben Sie keinen Zugriff … Authentifizierungsflow“).
- Vorher im Microsoft 365 Admin Center unter *Einstellungen → Organisationseinstellungen → Berichte*
  die Option „Anonymisierte Benutzer-, Gruppen- und Websitenamen anzeigen“ **deaktivieren**.
  Sonst können die Nutzungsdaten den Benutzern nicht zugeordnet werden.
- Optional `Install-Module ImportExcel -Scope CurrentUser` für eine zusätzliche .xlsx-Datei.
