<#
.SYNOPSIS
    Wertet einen Lizenzreport (Ordner von Get-LizenzReport.ps1) aus: Anmeldestatus, Entzugs-Kandidaten,
    Verteilung pro Amt und EmployeeType, Sparpotenzial. Erstellt CSVs, einen HTML-Bericht mit Grafiken
    und eine Zusammenfassung in der Konsole.

.DESCRIPTION
    Das Script braucht keine Module und aendert nichts. Es liest nur die CSV-Dateien des Reports.

    Eingabe:  Ordner LizenzReport_<Datum> mit 01_SKU-Uebersicht.csv und 02_Benutzer-Lizenzen.csv
    Ausgabe:  Unterordner Auswertung\ mit
              E5-Kandidaten.csv           alle Konten mit Massnahme (Entzug sofort / mit Rueckfrage / Eintritt)
              Kandidaten_pro_Amt.csv      Anzahl pro Amt, sofort / mit Rueckfrage -> an die Aemter verschicken
              Kandidaten_pro_EmployeeType.csv
              Anmeldestatus.csv           Verteilung aller E5-Lizenzen
              Auswertung.html             Bericht mit Grafiken (im Browser oeffnen, als PDF drucken)
              Zusammenfassung.txt         Text der Konsolenausgabe

    Kategorien (Anmeldestatus) und Massnahmen:
      Aktiv (Anmeldung <= InactiveDays)          Lizenz bleibt
      Neu, noch nie angemeldet (< InactiveDays)  Abwarten
      Vorbereiteter Eintritt                     Kein Entzug (deaktiviert, juenger als NewAccountDays, nie angemeldet)
      Inaktiv 91-180 Tage                        Mit Amt klaeren (Langzeitabwesenheit?)
      Inaktiv 181-365 Tage                       Mit Amt klaeren, Frist 14 Tage, dann entziehen
      Nie angemeldet (Konto > InactiveDays alt)  Mit Amt klaeren, Frist 14 Tage, dann entziehen
      Inaktiv > 365 Tage                         Sofort entziehen
      Konto deaktiviert (Austritt)               Sofort entziehen

.PARAMETER ReportPath
    Ordner des Reports (LizenzReport_<Datum>) oder direkt die Datei 02_Benutzer-Lizenzen.csv.
    Ohne Angabe wird der neueste Ordner LizenzReport_* im aktuellen Verzeichnis genommen.

.PARAMETER OutputPath
    Ausgabeordner. Standard: <ReportPath>\Auswertung

.PARAMETER InactiveDays
    Grenze fuer "inaktiv". Standard: 90 (muss zum Report passen)

.PARAMETER NewAccountDays
    Deaktivierte Konten juenger als diese Anzahl Tage ohne Anmeldung gelten als vorbereitete Eintritte. Standard: 60

.PARAMETER PreisE5
    Vertragspreis Microsoft 365 E5 pro Benutzer und Monat (CHF). Standard: 55 (Listenpreis, wird im Bericht als Annahme markiert)

.PARAMETER PreisP2
    Vertragspreis Entra ID P2 pro Benutzer und Monat (CHF). Standard: 8.50

.PARAMETER Ausnahmen
    Optionale CSV mit einer Spalte "UPN": Konten, die nie entzogen werden (Langzeitabwesenheit usw.).

.EXAMPLE
    .\New-LizenzAuswertung.ps1
    Wertet den neuesten LizenzReport_* im aktuellen Ordner aus.

.EXAMPLE
    .\New-LizenzAuswertung.ps1 -ReportPath C:\Temp\LizenzReport_2026-10-06 -PreisE5 38.20 -Ausnahmen .\Ausnahmen.csv
#>
[CmdletBinding()]
param(
    [string]$ReportPath,
    [string]$OutputPath,
    [int]$InactiveDays = 90,
    [int]$NewAccountDays = 60,
    [double]$PreisE5 = 55,
    [double]$PreisP2 = 8.5,
    [string]$Ausnahmen
)

$ErrorActionPreference = 'Stop'
$PreisIstAnnahme = -not $PSBoundParameters.ContainsKey('PreisE5')

# --------------------------------------------------------------------------------------------
# Hilfsfunktionen
# --------------------------------------------------------------------------------------------
function Write-Step([string]$Text) { Write-Host "==> $Text" -ForegroundColor Cyan }

function ConvertTo-Date($Value) {
    # Der Report schreibt Datumswerte je nach PowerShell-Sprache als ISO (2026-10-06T...) oder de-CH (06.10.2026 11:31:00)
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $s = ([string]$Value).Trim()
    $inv = [cultureinfo]::InvariantCulture
    $formats = @('dd.MM.yyyy HH:mm:ss', 'dd.MM.yyyy HH:mm', 'dd.MM.yyyy', 'yyyy-MM-ddTHH:mm:ssZ', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd')
    $out = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, [string[]]$formats, $inv, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$out)) { return $out }
    if ([datetime]::TryParse($s, $inv, [System.Globalization.DateTimeStyles]::None, [ref]$out)) { return $out }
    try { return [datetime]$s } catch { return $null }
}

function Format-Zahl($n) { return ('{0:N0}' -f [double]$n).Replace(',', "'").Replace('.', "'") }
function Format-Chf($n) { return 'CHF ' + (Format-Zahl ([math]::Round($n, 0))) }

function Get-Amt([string]$Ou) {
    # "Staff/Users/AIO/KTZGv2" -> AIO ; "AWA/VD/KTZG" -> AWA
    if ([string]::IsNullOrWhiteSpace($Ou)) { return '(keine OU)' }
    $p = $Ou -split '/'
    if ($p[0] -in @('Staff', 'Supplier', 'Anwaerter') -and $p.Count -ge 3) { return $p[2] }
    return $p[0]
}

function Export-Report($Data, [string]$Name) {
    $file = Join-Path $OutputPath "$Name.csv"
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $Data | Export-Csv -Path $file -NoTypeInformation -Delimiter ';' -Encoding $encoding
    Write-Host "    $file"
}

function Escape-Html([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }

function New-SvgBarChart([string]$Title, $Items, [string]$Subtitle = '') {
    # $Items: Objekte mit Label, Value, Color. Horizontales Balkendiagramm als Inline-SVG.
    $rowH = 30; $left = 290; $width = 900; $barMax = $width - $left - 90
    $max = ($Items | Measure-Object -Property Value -Maximum).Maximum
    if (-not $max) { $max = 1 }
    $height = 60 + $rowH * $Items.Count
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg xmlns='http://www.w3.org/2000/svg' width='$width' height='$height' font-family='Segoe UI, Arial, sans-serif' font-size='13'>")
    [void]$sb.Append("<text x='0' y='20' font-size='16' font-weight='bold' fill='#0b0b0b'>$(Escape-Html $Title)</text>")
    if ($Subtitle) { [void]$sb.Append("<text x='0' y='40' fill='#52514e' font-size='12'>$(Escape-Html $Subtitle)</text>") }
    $y = 55
    foreach ($it in $Items) {
        $w = [math]::Round($barMax * $it.Value / $max)
        if ($it.Value -gt 0 -and $w -lt 2) { $w = 2 }
        [void]$sb.Append("<text x='$($left - 10)' y='$($y + 19)' text-anchor='end' fill='#52514e'>$(Escape-Html $it.Label)</text>")
        [void]$sb.Append("<rect x='$left' y='$($y + 5)' width='$w' height='20' rx='3' fill='$($it.Color)'/>")
        [void]$sb.Append("<text x='$($left + $w + 6)' y='$($y + 19)' fill='#0b0b0b'>$(Format-Zahl $it.Value)</text>")
        $y += $rowH
    }
    [void]$sb.Append('</svg>')
    return $sb.ToString()
}

function New-HtmlTable($Rows, [string[]]$Columns) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table><thead><tr>')
    foreach ($c in $Columns) { [void]$sb.Append("<th>$(Escape-Html $c)</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($r in $Rows) {
        [void]$sb.Append('<tr>')
        foreach ($c in $Columns) {
            $v = $r.$c
            $cls = if ($v -is [int] -or $v -is [double] -or $v -is [long]) { " class='num'" } else { '' }
            if ($v -is [int] -or $v -is [long]) { $v = Format-Zahl $v }
            [void]$sb.Append("<td$cls>$(Escape-Html ([string]$v))</td>")
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table>')
    return $sb.ToString()
}

# --------------------------------------------------------------------------------------------
# 1. Eingabe finden und laden
# --------------------------------------------------------------------------------------------
Write-Step 'Lade Report'
if (-not $ReportPath) {
    $latest = Get-ChildItem -Directory -Filter 'LizenzReport_*' | Sort-Object Name -Descending | Select-Object -First 1
    if (-not $latest) { throw 'Kein Ordner LizenzReport_* gefunden. -ReportPath angeben.' }
    $ReportPath = $latest.FullName
}
if (Test-Path $ReportPath -PathType Leaf) { $ReportPath = Split-Path $ReportPath -Parent }
$ReportPath = (Resolve-Path $ReportPath).Path
$userFile = Join-Path $ReportPath '02_Benutzer-Lizenzen.csv'
$skuFile  = Join-Path $ReportPath '01_SKU-Uebersicht.csv'
if (-not (Test-Path $userFile)) { throw "Datei nicht gefunden: $userFile" }
if (-not $OutputPath) { $OutputPath = Join-Path $ReportPath 'Auswertung' }
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

$users = Import-Csv -Path $userFile -Delimiter ';' -Encoding UTF8
$skus  = if (Test-Path $skuFile) { Import-Csv -Path $skuFile -Delimiter ';' -Encoding UTF8 } else { @() }

# Stichtag aus dem Ordnernamen (LizenzReport_2026-10-06), sonst heute
$stichtag = Get-Date
if ((Split-Path $ReportPath -Leaf) -match '(\d{4}-\d{2}-\d{2})') { $stichtag = [datetime]$Matches[1] }
Write-Host "    $ReportPath"
Write-Host "    $($users.Count) Benutzer, Stichtag $($stichtag.ToString('dd.MM.yyyy'))"

$ausnahmeUpns = @{}
if ($Ausnahmen) {
    foreach ($a in (Import-Csv -Path $Ausnahmen -Delimiter ';')) {
        if ($a.UPN) { $ausnahmeUpns[([string]$a.UPN).ToLowerInvariant()] = $true }
    }
    Write-Host "    $($ausnahmeUpns.Count) Ausnahmen geladen"
}

# --------------------------------------------------------------------------------------------
# 2. Anmeldestatus und Massnahme pro E5-Benutzer
# --------------------------------------------------------------------------------------------
Write-Step 'Berechne Anmeldestatus'
$e5 = @($users | Where-Object { $_.Hat_E5 -eq 'True' })
if (-not $e5) { throw 'Keine E5-Benutzer im Report (Spalte Hat_E5).' }
$hasSignIn = [bool]($e5 | Where-Object { $_.Letzte_Anmeldung } | Select-Object -First 1)
if (-not $hasSignIn) {
    Write-Warning 'Der Report enthaelt keine Anmeldedaten (Letzte_Anmeldung leer). Entzugs-Kandidaten koennen nicht bestimmt werden.'
}

$Kategorien = [ordered]@{
    'Aktiv'      = @{ Label = "Aktiv (Anmeldung <= $InactiveDays Tage)";              Massnahme = 'Lizenz bleibt';                                        Gruppe = 'bleibt';     Color = '#2a78d6' }
    'Neu'        = @{ Label = "Neu (< $InactiveDays Tage), noch nie angemeldet";      Massnahme = 'Abwarten';                                             Gruppe = 'bleibt';     Color = '#9ec1ee' }
    'Eintritt'   = @{ Label = 'Vorbereiteter Eintritt (kein Entzug)';                 Massnahme = 'Kein Entzug - Lizenz erst ab erstem Arbeitstag';       Gruppe = 'eintritt';   Color = '#9ec1ee' }
    'Ausnahme'   = @{ Label = 'Ausnahmeliste (kein Entzug)';                          Massnahme = 'Kein Entzug - auf Ausnahmeliste';                      Gruppe = 'bleibt';     Color = '#9ec1ee' }
    'Inaktiv90'  = @{ Label = "Inaktiv $($InactiveDays + 1)-180 Tage";                Massnahme = 'Mit Amt klaeren (Langzeitabwesenheit?)';               Gruppe = 'rueckfrage'; Color = '#f2c27a' }
    'Inaktiv180' = @{ Label = 'Inaktiv 181-365 Tage';                                 Massnahme = 'Mit Amt klaeren, Frist 14 Tage, dann E5 entfernen';    Gruppe = 'rueckfrage'; Color = '#eda100' }
    'Nie'        = @{ Label = "Nie angemeldet (Konto > $InactiveDays Tage alt)";      Massnahme = 'Mit Amt klaeren, Frist 14 Tage, dann E5 entfernen';    Gruppe = 'rueckfrage'; Color = '#d98100' }
    'Inaktiv365' = @{ Label = 'Inaktiv > 365 Tage';                                   Massnahme = 'E5 entfernen (sofort)';                                Gruppe = 'sofort';     Color = '#b86300' }
    'Austritt'   = @{ Label = 'Konto deaktiviert (Austritt)';                         Massnahme = 'E5 entfernen (sofort)';                                Gruppe = 'sofort';     Color = '#8a4a00' }
}

foreach ($u in $e5) {
    $created  = ConvertTo-Date $u.Erstellt
    $lastSign = ConvertTo-Date $u.Letzte_Anmeldung
    $days = if ($u.Tage_seit_Anmeldung -match '^\d+$') { [int]$u.Tage_seit_Anmeldung } elseif ($lastSign) { [int]($stichtag - $lastSign).TotalDays } else { $null }
    # Falls das Datum nicht lesbar war, aber eine Anmeldung vorhanden ist: Tage-Spalte gilt
    if ($null -eq $lastSign -and $null -ne $days) { $lastSign = $stichtag.AddDays(-$days) }
    $enabled = $u.Aktiviert -eq 'True'
    $isNew   = $created -and $created -ge $stichtag.AddDays(-$NewAccountDays)
    $isOld   = $created -and $created -lt $stichtag.AddDays(-$InactiveDays)

    $key = if ($ausnahmeUpns.ContainsKey(([string]$u.UPN).ToLowerInvariant())) { 'Ausnahme' }
           elseif (-not $enabled -and $isNew -and $null -eq $lastSign) { 'Eintritt' }
           elseif (-not $enabled) { 'Austritt' }
           elseif (-not $hasSignIn) { 'Aktiv' }     # ohne Anmeldedaten keine Aussage -> wie aktiv behandeln
           elseif ($null -eq $lastSign -and $isOld) { 'Nie' }
           elseif ($null -eq $lastSign) { 'Neu' }
           elseif ($days -gt 365) { 'Inaktiv365' }
           elseif ($days -gt 180) { 'Inaktiv180' }
           elseif ($days -gt $InactiveDays) { 'Inaktiv90' }
           else { 'Aktiv' }

    $u | Add-Member -NotePropertyName _Key -NotePropertyValue $key -Force
    $u | Add-Member -NotePropertyName Kategorie -NotePropertyValue $Kategorien[$key].Label -Force
    $u | Add-Member -NotePropertyName Massnahme -NotePropertyValue $Kategorien[$key].Massnahme -Force
    $u | Add-Member -NotePropertyName Gruppe -NotePropertyValue $Kategorien[$key].Gruppe -Force
    $u | Add-Member -NotePropertyName Amt -NotePropertyValue (Get-Amt $u.AD_OU) -Force
    $u | Add-Member -NotePropertyName Tage -NotePropertyValue $days -Force
    if (-not $u.EmployeeType) { $u.EmployeeType = '(ohne EmployeeType)' }
}

$kandidaten = @($e5 | Where-Object { $_.Gruppe -in @('sofort', 'rueckfrage') })
$sofort     = @($kandidaten | Where-Object { $_.Gruppe -eq 'sofort' })
$rueckfrage = @($kandidaten | Where-Object { $_.Gruppe -eq 'rueckfrage' })
$eintritte  = @($e5 | Where-Object { $_.Gruppe -eq 'eintritt' })
$keyOrder   = @($Kategorien.Keys)

# --------------------------------------------------------------------------------------------
# 3. Tabellen
# --------------------------------------------------------------------------------------------
Write-Step 'Erstelle Tabellen'
$statusRows = foreach ($k in $keyOrder) {
    $n = @($e5 | Where-Object { $_._Key -eq $k }).Count
    if ($n -eq 0 -and $k -eq 'Ausnahme') { continue }
    [pscustomobject]@{ Anmeldestatus = $Kategorien[$k].Label; Anzahl = $n; Anteil = ('{0:P1}' -f ($n / $e5.Count)); Massnahme = $Kategorien[$k].Massnahme }
}

$amtRows = $kandidaten | Group-Object Amt | ForEach-Object {
    [pscustomobject]@{
        Amt               = $_.Name
        Sofort            = @($_.Group | Where-Object Gruppe -eq 'sofort').Count
        Mit_Rueckfrage    = @($_.Group | Where-Object Gruppe -eq 'rueckfrage').Count
        Total             = $_.Count
        E5_im_Amt         = @($e5 | Where-Object Amt -eq $_.Name).Count
    }
} | Sort-Object Total -Descending

$etRows = $kandidaten | Group-Object EmployeeType | ForEach-Object {
    $total = @($e5 | Where-Object EmployeeType -eq $_.Name).Count
    [pscustomobject]@{
        EmployeeType   = $_.Name
        Sofort         = @($_.Group | Where-Object Gruppe -eq 'sofort').Count
        Mit_Rueckfrage = @($_.Group | Where-Object Gruppe -eq 'rueckfrage').Count
        Kandidaten     = $_.Count
        E5_total       = $total
        Anteil         = ('{0:P0}' -f ($_.Count / $total))
    }
} | Sort-Object Kandidaten -Descending

$e5EtRows = $e5 | Group-Object EmployeeType | Sort-Object Count -Descending |
    ForEach-Object { [pscustomobject]@{ EmployeeType = $_.Name; E5 = $_.Count } }

$sortIdx = @{}; $i = 0; foreach ($k in @('Austritt', 'Inaktiv365', 'Nie', 'Inaktiv180', 'Inaktiv90', 'Eintritt', 'Ausnahme')) { $sortIdx[$k] = $i++ }
$kandidatenListe = @($kandidaten + $eintritte) | Sort-Object { $sortIdx[$_._Key] }, EmployeeType, Name |
    Select-Object Kategorie, Massnahme, Name, UPN, EmployeeType, Amt, AD_OU,
        @{ n = 'Erstellt'; e = { ([string]$_.Erstellt).Substring(0, [math]::Min(10, ([string]$_.Erstellt).Length)) } },
        @{ n = 'Letzte_Anmeldung'; e = { ([string]$_.Letzte_Anmeldung).Substring(0, [math]::Min(10, ([string]$_.Letzte_Anmeldung).Length)) } },
        @{ n = 'Tage_seit_Anmeldung'; e = { $_.Tage } }, Lizenzen, Zuweisung

$doppel = @($e5 | Where-Object { $_.Begruendung -like '*Doppellizenz*' })
$fehler = @($users | Where-Object { $_.Zuweisungsfehler })
$direkt = @($e5 | Where-Object { $_.Zuweisung -ne 'Gruppe' })

Export-Report $kandidatenListe 'E5-Kandidaten'
Export-Report $amtRows 'Kandidaten_pro_Amt'
Export-Report $etRows 'Kandidaten_pro_EmployeeType'
Export-Report $statusRows 'Anmeldestatus'

# --------------------------------------------------------------------------------------------
# 4. Sparpotenzial
# --------------------------------------------------------------------------------------------
$ersparnisSofort = $sofort.Count * $PreisE5 * 12
$ersparnisTotal  = $kandidaten.Count * $PreisE5 * 12
$preisHinweis = if ($PreisIstAnnahme) { " (Annahme Listenpreis CHF $PreisE5/Monat; Vertragspreis mit -PreisE5 angeben)" } else { " (Vertragspreis CHF $PreisE5/Monat)" }

# --------------------------------------------------------------------------------------------
# 5. HTML-Bericht
# --------------------------------------------------------------------------------------------
Write-Step 'Erstelle HTML-Bericht'
$chartStatus = New-SvgBarChart "E5-Lizenzen nach Anmeldestatus ($(Format-Zahl $e5.Count) Benutzer, $($stichtag.ToString('dd.MM.yyyy')))" `
    @(foreach ($k in $keyOrder) { $n = @($e5 | Where-Object { $_._Key -eq $k }).Count; if ($n -or $k -ne 'Ausnahme') { [pscustomobject]@{ Label = $Kategorien[$k].Label; Value = $n; Color = $Kategorien[$k].Color } } }) `
    "Blau = Lizenz bleibt | Orange = Kandidaten mit Rueckfrage ($($rueckfrage.Count)) | Braun = Sofortmassnahmen ($($sofort.Count))"
$chartEt  = New-SvgBarChart "Entzugs-Kandidaten nach EmployeeType ($($kandidaten.Count))" @($etRows | ForEach-Object { [pscustomobject]@{ Label = $_.EmployeeType; Value = $_.Kandidaten; Color = '#eda100' } })
$chartAmt = New-SvgBarChart "Entzugs-Kandidaten nach Amt (Top 15 von $($amtRows.Count))" @($amtRows | Select-Object -First 15 | ForEach-Object { [pscustomobject]@{ Label = $_.Amt; Value = $_.Total; Color = '#eda100' } })
$chartE5  = New-SvgBarChart "Alle E5-Lizenzen nach EmployeeType" @($e5EtRows | ForEach-Object { [pscustomobject]@{ Label = $_.EmployeeType; Value = $_.E5; Color = '#2a78d6' } })

$skuRows = $skus | ForEach-Object { [pscustomobject]@{ Produkt = $_.Produkt; Gekauft = [int]$_.Gekauft; Zugewiesen = [int]$_.Zugewiesen; Frei = [int]$_.Frei } } |
    Where-Object { $_.Gekauft -lt 100000 } | Sort-Object Zugewiesen -Descending

$html = @"
<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8">
<title>Lizenzauswertung $($stichtag.ToString('dd.MM.yyyy'))</title>
<style>
 body { font-family: 'Segoe UI', Arial, sans-serif; color: #0b0b0b; background: #fcfcfb; max-width: 1000px; margin: 30px auto; padding: 0 20px; line-height: 1.45; }
 h1 { font-size: 24px; } h2 { font-size: 18px; margin-top: 36px; border-bottom: 1px solid #e6e5e0; padding-bottom: 4px; }
 table { border-collapse: collapse; margin: 12px 0; font-size: 13px; } th, td { border-bottom: 1px solid #e6e5e0; padding: 5px 10px; text-align: left; }
 th { background: #f1f0eb; } td.num { text-align: right; font-variant-numeric: tabular-nums; }
 .kpi { display: inline-block; min-width: 160px; margin: 8px 16px 8px 0; padding: 12px 16px; border: 1px solid #e6e5e0; border-radius: 6px; background: #fff; }
 .kpi b { display: block; font-size: 26px; } .kpi span { color: #52514e; font-size: 12px; }
 .hinweis { background: #fff6e0; border-left: 4px solid #eda100; padding: 8px 12px; margin: 12px 0; }
 .muted { color: #52514e; font-size: 12px; } svg { display: block; margin: 12px 0 24px; }
 @media print { body { max-width: none; } h2 { page-break-after: avoid; } }
</style></head><body>
<h1>Lizenzauswertung Microsoft 365 – Stand $($stichtag.ToString('dd.MM.yyyy'))</h1>
<p class="muted">Quelle: $(Escape-Html $ReportPath) · Erstellt $((Get-Date).ToString('dd.MM.yyyy HH:mm')) mit New-LizenzAuswertung.ps1 · Grenzen: inaktiv ab $InactiveDays Tagen, Eintritte bis $NewAccountDays Tage</p>

<div class="kpi"><b>$(Format-Zahl $e5.Count)</b><span>E5-Lizenzen zugewiesen</span></div>
<div class="kpi"><b>$(Format-Zahl $kandidaten.Count)</b><span>Entzugs-Kandidaten ($('{0:P1}' -f ($kandidaten.Count / $e5.Count)))</span></div>
<div class="kpi"><b>$(Format-Zahl $sofort.Count)</b><span>davon sofort</span></div>
<div class="kpi"><b>$(Format-Zahl $eintritte.Count)</b><span>vorbereitete Eintritte (kein Entzug)</span></div>
<div class="kpi"><b>$(Format-Chf $ersparnisTotal)</b><span>Ersparnis pro Jahr bei Entzug aller Kandidaten$(if ($PreisIstAnnahme) { ' (Annahme)' })</span></div>

$(if (-not $hasSignIn) { '<div class="hinweis"><b>Keine Anmeldedaten im Report.</b> Entzugs-Kandidaten konnten nicht bestimmt werden. Report mit aktivierter Rolle Security Administrator wiederholen.</div>' })

<h2>1. Lizenzbestand</h2>
$(New-HtmlTable $skuRows @('Produkt', 'Gekauft', 'Zugewiesen', 'Frei'))

<h2>2. E5-Lizenzen nach Anmeldestatus</h2>
$chartStatus
$(New-HtmlTable $statusRows @('Anmeldestatus', 'Anzahl', 'Anteil', 'Massnahme'))
<p class="muted">Grundlage ist das Feld signInActivity in Entra ID: der juengere der beiden Zeitstempel "letzte interaktive Anmeldung" (Person tippt Passwort/MFA) und "letzte nicht-interaktive Anmeldung" (Geraet erneuert im Hintergrund ein Token). "Inaktiv" heisst: weder die Person noch eines ihrer Geraete hat sich in dieser Zeit bei Microsoft 365 oder Entra ID gemeldet.</p>

<h2>3. Kandidaten nach EmployeeType</h2>
$chartEt
$(New-HtmlTable $etRows @('EmployeeType', 'Sofort', 'Mit_Rueckfrage', 'Kandidaten', 'E5_total', 'Anteil'))

<h2>4. Kandidaten nach Amt</h2>
$chartAmt
$(New-HtmlTable $amtRows @('Amt', 'Sofort', 'Mit_Rueckfrage', 'Total', 'E5_im_Amt'))
<p class="muted">Vollstaendige Liste mit Namen: E5-Kandidaten.csv. Pro Amt: Kandidaten_pro_Amt.csv.</p>

<h2>5. Alle E5-Lizenzen nach EmployeeType</h2>
$chartE5

<h2>6. Weitere Befunde</h2>
<ul>
 <li><b>$($eintritte.Count) vorbereitete Eintritte</b> (deaktiviert, juenger als $NewAccountDays Tage, nie angemeldet) belegen eine E5, bevor die Person anfaengt. Massnahme: Lizenzgruppe mit Bedingung <code>accountEnabled = true</code>.</li>
 <li><b>$($doppel.Count) Doppellizenzen</b> (E5 plus P1/P2/Power BI Pro, bereits in E5 enthalten).</li>
 <li><b>$($fehler.Count) Lizenzzuweisungsfehler</b> (Spalte Zuweisungsfehler in 02_Benutzer-Lizenzen.csv).</li>
 <li><b>$($direkt.Count) E5 nicht rein ueber Gruppen</b> zugewiesen (direkt oder gemischt). Ziel: 0.</li>
</ul>

<h2>7. Sparpotenzial</h2>
<table><thead><tr><th>Massnahme</th><th>Lizenzen</th><th>Ersparnis pro Jahr</th></tr></thead><tbody>
<tr><td>Sofort (Austritt, inaktiv &gt; 365 Tage)</td><td class="num">$(Format-Zahl $sofort.Count)</td><td class="num">$(Format-Chf $ersparnisSofort)</td></tr>
<tr><td>Nach Rueckfrage bei den Aemtern</td><td class="num">$(Format-Zahl $rueckfrage.Count)</td><td class="num">$(Format-Chf ($ersparnisTotal - $ersparnisSofort))</td></tr>
<tr><td><b>Total</b></td><td class="num"><b>$(Format-Zahl $kandidaten.Count)</b></td><td class="num"><b>$(Format-Chf $ersparnisTotal)</b></td></tr>
</tbody></table>
<p class="muted">Preis E5$(Escape-Html $preisHinweis). Entra ID P2: CHF $PreisP2/Monat. Umstellung E5 -&gt; P2 spart CHF $(Format-Zahl (($PreisE5 - $PreisP2) * 12)) pro Benutzer und Jahr.</p>

<h2>8. Naechste Schritte</h2>
<ol>
 <li>Sofortmassnahmen ($($sofort.Count)): Konten aus der E5-Lizenzgruppe entfernen. Vorher Postfach/OneDrive pruefen, Postfach ggf. in freigegebenes Postfach umwandeln.</li>
 <li>Kandidaten_pro_Amt.csv an die Aemter, Frist 14 Tage, dann Entzug der uebrigen $($rueckfrage.Count).</li>
 <li>accountEnabled-Bedingung in die Lizenzgruppe aufnehmen (Eintritte).</li>
 <li>Nutzungsberichte beschaffen (Rolle Reports Reader, Admin Consent Reports.Read.All) fuer Block 2: aktive E5 ohne Office-/Mail-Nutzung -&gt; Entra ID P2.</li>
 <li>Report quartalsweise wiederholen.</li>
</ol>
</body></html>
"@
$htmlFile = Join-Path $OutputPath 'Auswertung.html'
[IO.File]::WriteAllText($htmlFile, $html, [Text.Encoding]::UTF8)
Write-Host "    $htmlFile"

# --------------------------------------------------------------------------------------------
# 6. Zusammenfassung
# --------------------------------------------------------------------------------------------
Write-Step 'Zusammenfassung'
$summary = New-Object System.Collections.Generic.List[string]
$summary.Add("Lizenzauswertung Stand $($stichtag.ToString('dd.MM.yyyy'))  |  Report: $ReportPath")
$summary.Add("E5 zugewiesen: $(Format-Zahl $e5.Count)   Kandidaten: $($kandidaten.Count) (sofort $($sofort.Count), Rueckfrage $($rueckfrage.Count))   Eintritte: $($eintritte.Count)")
$summary.Add("Ersparnis/Jahr: sofort $(Format-Chf $ersparnisSofort), total $(Format-Chf $ersparnisTotal)$preisHinweis")
$summary.Add('')
$summary.Add(($statusRows | Format-Table Anmeldestatus, Anzahl, Anteil, Massnahme -AutoSize | Out-String).Trim())
$summary.Add('')
$summary.Add('Kandidaten nach EmployeeType:')
$summary.Add(($etRows | Format-Table EmployeeType, Sofort, Mit_Rueckfrage, Kandidaten, E5_total, Anteil -AutoSize | Out-String).Trim())
$summary.Add('')
$summary.Add('Kandidaten nach Amt (Top 15):')
$summary.Add(($amtRows | Select-Object -First 15 | Format-Table Amt, Sofort, Mit_Rueckfrage, Total, E5_im_Amt -AutoSize | Out-String).Trim())
$summary.Add('')
$summary.Add("Doppellizenzen: $($doppel.Count)   Zuweisungsfehler: $($fehler.Count)   E5 nicht rein ueber Gruppe: $($direkt.Count)")
$summary.Add("Ausgabe: $OutputPath")
$summaryText = $summary -join [Environment]::NewLine
[IO.File]::WriteAllText((Join-Path $OutputPath 'Zusammenfassung.txt'), $summaryText, [Text.Encoding]::UTF8)
Write-Host $summaryText
Write-Host "Fertig. Auswertung.html im Browser oeffnen (dort auch als PDF druckbar)." -ForegroundColor Green
