<#
.SYNOPSIS
    Liest alle Microsoft-365-/Entra-Lizenzen inkl. Nutzungsdaten aus und markiert
    Kandidaten, die von Microsoft 365 E5 auf Entra ID P2 (oder gar keine Lizenz)
    umgestellt werden koennen.

.DESCRIPTION
    Das Script ist read-only - es wird KEINE Lizenz veraendert.

    Es erstellt im Ausgabeordner:
      01_SKU-Uebersicht.csv      Gekaufte vs. zugewiesene Lizenzen pro Produkt
      02_Benutzer-Lizenzen.csv   Ein Eintrag pro Benutzer mit Lizenzen, Nutzung und Empfehlung
      03_E5-Kandidaten.csv       Nur die E5-Benutzer mit einer Empfehlung != "E5 behalten"
      Rohdaten\*.csv             Originale Microsoft-365-Nutzungsberichte

    Anmeldung: OHNE Zusatzmodule. Das Script oeffnet den Browser, meldet sich wie Azure PowerShell
    (Microsoft-eigene App, keine Admin-Freigabe noetig) per OAuth an und liest die Daten direkt
    per REST von Microsoft Graph. Es werden weder Microsoft.Graph noch Az benoetigt.

    Datenquellen (Microsoft Graph REST):
      - subscribedSkus                      Lizenzbestand
      - users (+ signInActivity)            Lizenzen, Zuweisung (direkt/Gruppe), letzte Anmeldung, AD-OU
      - getOffice365ActiveUserDetail        Letzte Aktivitaet Exchange / OneDrive / SharePoint / Teams
      - getOffice365ActivationsUserDetail   Office-Desktop-Apps auf Windows/Mac aktiviert?
      - getTeamsUserActivityUserDetail      Anrufe / Meetings in Teams
      - optional: Teams-Telefonie (Get-CsOnlineUser), da Teams Phone ein E5-Merkmal ist

.PARAMETER OutputPath
    Ausgabeordner. Standard: .\LizenzReport_<Datum>

.PARAMETER InactiveDays
    Ab wie vielen Tagen ohne Anmeldung ein Konto als inaktiv gilt. Standard: 90

.PARAMETER Period
    Zeitraum der Nutzungsberichte (D7, D30, D90, D180). Standard: D90

.PARAMETER UseDeviceCode
    Anmeldung per Geraetecode (Code auf https://microsoft.com/devicelogin eingeben).
    Hilft, wenn das Anmeldefenster nicht erscheint, z. B. in der PowerShell ISE oder auf einem Server.

.PARAMETER TenantId
    Tenant-ID oder Domain (z. B. zg.ch), falls das Konto in mehreren Tenants existiert.

.PARAMETER RequestScopes
    Fordert die benoetigten Berechtigungen bei der Anmeldung explizit an (Zustimmungsdialog).
    Standard: nur die bereits freigegebenen Berechtigungen verwenden (kein Dialog).

.PARAMETER IncludeTeamsPhone
    Prueft zusaetzlich mit dem MicrosoftTeams-Modul, wer Teams-Telefonie (Enterprise Voice) nutzt.

.EXAMPLE
    .\Get-LizenzReport.ps1

.EXAMPLE
    .\Get-LizenzReport.ps1 -InactiveDays 60 -Period D180 -IncludeTeamsPhone

.NOTES
    Voraussetzungen:
      keine - nur Windows PowerShell 5.1 oder PowerShell 7
      (optional) Install-Module MicrosoftTeams -Scope CurrentUser
      (optional) Install-Module ImportExcel -Scope CurrentUser   -> zusaetzlich eine .xlsx-Datei

    Rollen: Global Reader + Reports Reader reichen aus.

    WICHTIG: Im Microsoft 365 Admin Center unter
      Einstellungen > Organisationseinstellungen > Berichte
    muss "Anonymisierte Benutzer-, Gruppen- und Websitenamen in allen Berichten anzeigen"
    DEAKTIVIERT sein, sonst sind die Benutzernamen in den Nutzungsberichten verschluesselt
    und koennen nicht zugeordnet werden. Das Script prueft das und warnt.
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path (Get-Location) ("LizenzReport_{0:yyyy-MM-dd}" -f (Get-Date))),
    [int]$InactiveDays = 90,
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D90',
    [switch]$IncludeTeamsPhone,
    [switch]$UseDeviceCode,
    [string]$TenantId,
    [switch]$RequestScopes
)

$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------------------------
# Lizenz-Kategorien (SkuPartNumber, Regex). Bei Bedarf an eure Vertraege anpassen.
# --------------------------------------------------------------------------------------------
$SkuPattern = @{
    E5       = '^(SPE_E5|SPE_E5_NOPSTNCONF|Microsoft_365_E5.*|ENTERPRISEPREMIUM|ENTERPRISEPREMIUM_NOPSTNCONF|Office_365_E5.*|M365_E5.*)$'
    E3       = '^(SPE_E3|Microsoft_365_E3.*|ENTERPRISEPACK|Office_365_E3.*|O365_w/o_Teams_Bundle_M3)$'
    F3       = '^(SPE_F1|M365_F1|DESKLESSPACK|SPE_F5_.*)$'
    EntraP2  = '^(AAD_PREMIUM_P2|Microsoft_Entra_ID_P2)$'
    EntraP1  = '^(AAD_PREMIUM)$'
    EMS      = '^(EMS|EMSPREMIUM)$'
    PowerBI  = '^(POWER_BI_PRO)$'
    LTSC     = 'LTSC'
}

# Fallback-Namen, falls die offizielle Microsoft-Liste nicht geladen werden kann
$SkuNameFallback = @{
    'SPE_E5'             = 'Microsoft 365 E5'
    'SPE_E5_NOPSTNCONF'  = 'Microsoft 365 E5 (ohne Audio Conferencing)'
    'SPE_E3'             = 'Microsoft 365 E3'
    'SPE_F1'             = 'Microsoft 365 F3'
    'ENTERPRISEPREMIUM'  = 'Office 365 E5'
    'ENTERPRISEPACK'     = 'Office 365 E3'
    'DESKLESSPACK'       = 'Office 365 F3'
    'AAD_PREMIUM'        = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'     = 'Microsoft Entra ID P2'
    'EMS'                = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'         = 'Enterprise Mobility + Security E5'
    'POWER_BI_PRO'       = 'Power BI Pro'
    'POWER_BI_STANDARD'  = 'Power BI (free)'
    'FLOW_FREE'          = 'Power Automate (free)'
    'MCOEV'              = 'Teams Phone Standard'
    'MCOMEETADV'         = 'Teams Audio Conferencing'
    'STREAM'             = 'Microsoft Stream'
    'Teams_Ess'          = 'Teams Essentials'
}

# --------------------------------------------------------------------------------------------
# Hilfsfunktionen
# --------------------------------------------------------------------------------------------
function Write-Step([string]$Text) { Write-Host "==> $Text" -ForegroundColor Cyan }

function ConvertTo-Date($Value) {
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    try { return [datetime]$Value } catch { return $null }
}

function Get-DaysSince($Date) {
    if ($null -eq $Date) { return $null }
    return [int]((Get-Date) - $Date).TotalDays
}

function Get-OuFromDn([string]$Dn) {
    if ([string]::IsNullOrWhiteSpace($Dn)) { return $null }
    # "CN=Muster Max,OU=Extern,OU=Benutzer,DC=zg,DC=ch" -> "Extern/Benutzer"
    $ous = ($Dn -split '(?<!\\),') | Where-Object { $_ -like 'OU=*' } | ForEach-Object { $_.Substring(3) }
    return ($ous -join '/')
}

$Graph = 'https://graph.microsoft.com/v1.0'
$script:GraphHeaders = $null

# --------------------------------------------------------------------------------------------
# Anmeldung (OAuth 2.0 Authorization Code + PKCE, ohne Zusatzmodule)
# --------------------------------------------------------------------------------------------
# Client-ID von "Microsoft Azure PowerShell" (Microsoft-eigene oeffentliche App, erlaubt http://localhost)
$script:ClientId = '1950a258-227b-4e31-a9cf-717495945fc2'
$script:Auth = $null   # access_token, refresh_token, expires (DateTime)

function ConvertTo-Base64Url([byte[]]$Bytes) {
    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-AuthScopes {
    if ($RequestScopes) {
        return 'https://graph.microsoft.com/User.Read.All https://graph.microsoft.com/Directory.Read.All ' +
               'https://graph.microsoft.com/Organization.Read.All https://graph.microsoft.com/AuditLog.Read.All ' +
               'https://graph.microsoft.com/Reports.Read.All https://graph.microsoft.com/ReportSettings.Read.All offline_access'
    }
    return 'https://graph.microsoft.com/.default offline_access'
}

function Set-AuthFromTokenResponse($Response) {
    $script:Auth = @{
        access_token  = $Response.access_token
        refresh_token = $Response.refresh_token
        expires       = (Get-Date).AddSeconds([int]$Response.expires_in - 120)
    }
    $script:GraphHeaders = @{ Authorization = "Bearer $($Response.access_token)"; ConsistencyLevel = 'eventual' }
}

function Invoke-TokenEndpoint([hashtable]$Body) {
    $tenant = if ($TenantId) { $TenantId } else { 'organizations' }
    try {
        return Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token" `
            -ContentType 'application/x-www-form-urlencoded' -Body $Body -ErrorAction Stop
    }
    catch {
        $detail = $_.ErrorDetails.Message
        if (-not $detail -and $_.Exception.Response) {
            try { $detail = (New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() } catch {}
        }
        if ($detail) {
            try { $j = $detail | ConvertFrom-Json; $detail = "$($j.error): $($j.error_description)" } catch {}
        }
        throw "Token-Anfrage fehlgeschlagen: $detail $($_.Exception.Message)"
    }
}

function Connect-GraphBrowser {
    # Lokaler Listener auf 127.0.0.1 (kein HttpListener -> keine Adminrechte / URL-ACL noetig)
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $port = $listener.LocalEndpoint.Port
    $redirect = "http://localhost:$port"

    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $buf = New-Object byte[] 32; $rng.GetBytes($buf); $verifier = ConvertTo-Base64Url $buf
    $rng.GetBytes($buf); $state = ConvertTo-Base64Url $buf
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $challenge = ConvertTo-Base64Url ($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier)))

    $tenant = if ($TenantId) { $TenantId } else { 'organizations' }
    $url = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/authorize?" +
        "client_id=$script:ClientId&response_type=code&redirect_uri=$([uri]::EscapeDataString($redirect))" +
        "&response_mode=query&scope=$([uri]::EscapeDataString((Get-AuthScopes)))" +
        "&state=$state&code_challenge=$challenge&code_challenge_method=S256&prompt=select_account"

    Write-Host '    Browser wird geoeffnet. Bitte dort anmelden ...'
    Write-Host "    Falls kein Fenster erscheint, diese Adresse im Browser oeffnen:`n    $url"
    try { Start-Process $url | Out-Null } catch { }

    $code = $null
    try {
        $deadline = (Get-Date).AddMinutes(5)
        while ($null -eq $code) {
            if ((Get-Date) -gt $deadline) { throw 'Keine Anmeldung innerhalb von 5 Minuten.' }
            if (-not $listener.Pending()) { Start-Sleep -Milliseconds 200; continue }
            $client = $listener.AcceptTcpClient()
            try {
                $stream = $client.GetStream()
                $reader = New-Object IO.StreamReader($stream)
                $requestLine = $reader.ReadLine()
                # restliche Header lesen und verwerfen
                while (($line = $reader.ReadLine()) -ne $null -and $line -ne '') { }
                $path = ($requestLine -split ' ')[1]
                $query = [System.Web.HttpUtility]::ParseQueryString(([uri]"http://localhost$path").Query)
                $msg = 'Anmeldung fehlgeschlagen. Dieses Fenster kann geschlossen werden.'
                if ($query['state'] -eq $state -and $query['code']) {
                    $code = $query['code']
                    $msg = 'Anmeldung erfolgreich. Dieses Fenster kann geschlossen werden.'
                }
                elseif ($query['error']) {
                    $msg = "Fehler: $($query['error']) - $($query['error_description'])"
                    Write-Warning $msg
                }
                $html = "<html><body style='font-family:sans-serif;padding:40px'><h2>Get-LizenzReport</h2><p>$msg</p></body></html>"
                $bytes = [Text.Encoding]::UTF8.GetBytes($html)
                $header = "HTTP/1.1 200 OK`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
                $hb = [Text.Encoding]::ASCII.GetBytes($header)
                $stream.Write($hb, 0, $hb.Length); $stream.Write($bytes, 0, $bytes.Length); $stream.Flush()
                if ($query['error']) { throw "Anmeldung abgelehnt: $($query['error']) - $($query['error_description'])" }
            }
            finally { $client.Close() }
        }
    }
    finally { $listener.Stop() }

    $resp = Invoke-TokenEndpoint @{
        client_id = $script:ClientId; grant_type = 'authorization_code'; code = $code
        redirect_uri = $redirect; code_verifier = $verifier; scope = (Get-AuthScopes)
    }
    Set-AuthFromTokenResponse $resp
}

function Connect-GraphDeviceCode {
    $tenant = if ($TenantId) { $TenantId } else { 'organizations' }
    $dc = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode" `
        -Body @{ client_id = $script:ClientId; scope = (Get-AuthScopes) } -ErrorAction Stop
    Write-Host "    $($dc.message)" -ForegroundColor Yellow
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds ([int]$dc.interval)
        try {
            $resp = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token" `
                -Body @{ client_id = $script:ClientId; grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; device_code = $dc.device_code } -ErrorAction Stop
            Set-AuthFromTokenResponse $resp
            return
        }
        catch {
            $detail = $_.ErrorDetails.Message
            if ($detail -match 'authorization_pending|slow_down') { continue }
            throw "Geraetecode-Anmeldung fehlgeschlagen: $detail $($_.Exception.Message)"
        }
    }
    throw 'Geraetecode abgelaufen.'
}

function Update-GraphToken {
    # Token vor Ablauf still erneuern (grosse Tenants brauchen laenger als 1 Stunde)
    if ($script:Auth -and (Get-Date) -ge $script:Auth.expires -and $script:Auth.refresh_token) {
        $resp = Invoke-TokenEndpoint @{
            client_id = $script:ClientId; grant_type = 'refresh_token'
            refresh_token = $script:Auth.refresh_token; scope = (Get-AuthScopes)
        }
        Set-AuthFromTokenResponse $resp
    }
}

function Get-TokenInfo {
    # Payload des Access-Tokens lesen (nur zur Anzeige: Benutzer, Tenant, Berechtigungen)
    try {
        $part = ($script:Auth.access_token -split '\.')[1].Replace('-', '+').Replace('_', '/')
        switch ($part.Length % 4) { 2 { $part += '==' } 3 { $part += '=' } }
        return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($part)) | ConvertFrom-Json
    }
    catch { return $null }
}

function Invoke-Graph([string]$Uri) {
    # Einzelne GET-Abfrage (JSON)
    Update-GraphToken
    if ($Uri -notmatch '^https://') { $Uri = "$Graph/$Uri" }
    return Invoke-RestMethod -Method GET -Uri $Uri -Headers $script:GraphHeaders -ErrorAction Stop
}

function Get-GraphCollection([string]$Uri) {
    # GET mit Paging (@odata.nextLink)
    $all = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $page = Invoke-Graph $next
        foreach ($v in $page.value) { $all.Add($v) }
        $next = $page.'@odata.nextLink'
        if ($all.Count -gt 0 -and $all.Count % 5000 -eq 0) { Write-Host "    ... $($all.Count)" }
    }
    return $all
}

function Save-GraphFile([string]$Uri, [string]$File) {
    # Berichte liefern ein 302 auf eine vorab signierte Download-URL. Die darf OHNE
    # Authorization-Header geladen werden, sonst lehnt der Speicher die Anfrage ab.
    Update-GraphToken
    $location = $null
    try {
        $r = Invoke-WebRequest -Method GET -Uri $Uri -Headers $script:GraphHeaders -MaximumRedirection 0 -UseBasicParsing -ErrorAction Stop
        if ($r.StatusCode -eq 200) { [IO.File]::WriteAllBytes($File, $r.Content); return }
        $location = $r.Headers['Location']
    }
    catch {
        $resp = $_.Exception.Response
        if ($null -eq $resp) { throw }
        $code = [int]$resp.StatusCode
        if ($code -lt 300 -or $code -ge 400) { throw }
        $location = if ($resp.Headers -is [System.Net.WebHeaderCollection]) { $resp.Headers['Location'] }
                    else { [string]$resp.Headers.Location }
    }
    if (-not $location) { throw 'Keine Download-URL erhalten' }
    Invoke-WebRequest -Uri $location -OutFile $File -UseBasicParsing -ErrorAction Stop
}

function Get-GraphReport([string]$Function, [string]$FileName) {
    $file = Join-Path $RawPath $FileName
    try {
        Save-GraphFile "$Graph/reports/$Function" $file
        $rows = Import-Csv -Path $file
        Write-Host "    $FileName : $($rows.Count) Zeilen"
        return $rows
    }
    catch {
        Write-Warning "Bericht $Function konnte nicht geladen werden: $($_.Exception.Message)"
        return @()
    }
}

function New-UpnIndex($Rows) {
    $index = @{}
    foreach ($row in $Rows) {
        $upn = [string]$row.'User Principal Name'
        if ($upn) { $index[$upn.ToLowerInvariant()] = $row }
    }
    return $index
}

function Export-Report($Data, [string]$Name) {
    $file = Join-Path $OutputPath "$Name.csv"
    # Semikolon + UTF-8 mit BOM -> oeffnet in Excel (de-CH) direkt korrekt
    # (Windows PowerShell 5.1 schreibt bei "UTF8" bereits ein BOM)
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $Data | Export-Csv -Path $file -NoTypeInformation -Delimiter ';' -Encoding $encoding
    Write-Host "    $file"
}

# --------------------------------------------------------------------------------------------
# 1. Verbinden
# --------------------------------------------------------------------------------------------
Write-Step 'Anmeldung bei Microsoft'
Write-Host "    PowerShell $($PSVersionTable.PSVersion), ohne Zusatzmodule"
# Windows PowerShell 5.1: TLS 1.2 erzwingen, sonst lehnt login.microsoftonline.com ab
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Web

if ($UseDeviceCode) { Connect-GraphDeviceCode } else { Connect-GraphBrowser }
$ti = Get-TokenInfo
if ($ti) {
    Write-Host "    Angemeldet als $($ti.upn), Tenant $($ti.tid)"
    Write-Host "    Berechtigungen: $($ti.scp)"
    foreach ($need in 'User.Read.All', 'AuditLog.Read.All', 'Reports.Read.All') {
        if ($ti.scp -notmatch [regex]::Escape($need)) {
            Write-Warning "Berechtigung $need fehlt im Token. Falls Daten fehlen: Script mit -RequestScopes starten."
        }
    }
}

New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
$RawPath = Join-Path $OutputPath 'Rohdaten'
New-Item -ItemType Directory -Path $RawPath -Force | Out-Null

try {
    $reportSettings = Invoke-Graph 'admin/reportSettings'
    if ($reportSettings.displayConcealedNames) {
        Write-Warning ('Die Nutzungsberichte sind anonymisiert (displayConcealedNames = true). ' +
            'Nutzungsdaten koennen den Benutzern NICHT zugeordnet werden. ' +
            'Im Admin Center unter Einstellungen > Organisationseinstellungen > Berichte deaktivieren und Script erneut starten.')
    }
}
catch {
    Write-Warning 'Konnte die Berichtseinstellungen nicht pruefen (ReportSettings.Read.All fehlt?).'
}

# --------------------------------------------------------------------------------------------
# 2. Lizenzbestand
# --------------------------------------------------------------------------------------------
Write-Step 'Lade Produktnamen und Lizenzbestand'
$SkuNames = @{} + $SkuNameFallback
try {
    $csvUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    $productCsv = Invoke-RestMethod -Uri $csvUrl -TimeoutSec 30 | ConvertFrom-Csv
    foreach ($p in $productCsv) { $SkuNames[$p.String_Id] = $p.Product_Display_Name }
}
catch {
    Write-Warning 'Offizielle Produktnamen-Liste nicht erreichbar, verwende interne Liste.'
}

$skus = Get-GraphCollection 'subscribedSkus'
$SkuById = @{}
foreach ($sku in $skus) { $SkuById[[string]$sku.SkuId] = $sku }

function Get-SkuName([string]$PartNumber) {
    if ($SkuNames.ContainsKey($PartNumber)) { return $SkuNames[$PartNumber] }
    return $PartNumber
}

$skuOverview = foreach ($sku in $skus) {
    [pscustomobject]@{
        Produkt       = Get-SkuName $sku.SkuPartNumber
        SkuPartNumber = $sku.SkuPartNumber
        Gekauft       = $sku.PrepaidUnits.Enabled
        Zugewiesen    = $sku.ConsumedUnits
        Frei          = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
        Gesperrt      = $sku.PrepaidUnits.Suspended
        Warnung       = $sku.PrepaidUnits.Warning
        Status        = $sku.CapabilityStatus
    }
}

# --------------------------------------------------------------------------------------------
# 3. Benutzer
# --------------------------------------------------------------------------------------------
Write-Step 'Lade Benutzer (kann bei vielen Konten einige Minuten dauern)'
$userProps = @(
    'id', 'displayName', 'userPrincipalName', 'mail', 'accountEnabled', 'userType',
    'department', 'jobTitle', 'companyName', 'employeeType', 'officeLocation',
    'createdDateTime', 'onPremisesSyncEnabled', 'onPremisesDistinguishedName',
    'assignedLicenses', 'licenseAssignmentStates', 'signInActivity'
)
$select = $userProps -join ','
try {
    $users = Get-GraphCollection "users?`$select=$select&`$top=999"
}
catch {
    # signInActivity braucht AuditLog.Read.All. Falls das fehlt: ohne letzte Anmeldung weiterarbeiten.
    Write-Warning "Letzte Anmeldung (signInActivity) nicht lesbar, lade Benutzer ohne dieses Feld. ($($_.Exception.Message))"
    $select = ($userProps | Where-Object { $_ -ne 'signInActivity' }) -join ','
    $users = Get-GraphCollection "users?`$select=$select&`$top=999"
}
Write-Host "    $($users.Count) Benutzer"

# --------------------------------------------------------------------------------------------
# 4. Nutzungsberichte
# --------------------------------------------------------------------------------------------
Write-Step "Lade Nutzungsberichte ($Period)"
$activeUsers = New-UpnIndex (Get-GraphReport "getOffice365ActiveUserDetail(period='$Period')" 'Office365ActiveUserDetail.csv')
# Aktivierungsbericht hat eine Zeile pro Benutzer UND Produkt -> Windows/Mac-Aktivierungen summieren
$activations = @{}
foreach ($row in (Get-GraphReport 'getOffice365ActivationsUserDetail' 'Office365ActivationsUserDetail.csv')) {
    $key = ([string]$row.'User Principal Name').ToLowerInvariant()
    if (-not $key) { continue }
    $activations[$key] = [int]$activations[$key] + [int]("0" + $row.Windows) + [int]("0" + $row.Mac)
}
$teamsUsage  = New-UpnIndex (Get-GraphReport "getTeamsUserActivityUserDetail(period='$Period')" 'TeamsUserActivityUserDetail.csv')

$voiceUsers = @{}
if ($IncludeTeamsPhone) {
    Write-Step 'Pruefe Teams-Telefonie (Enterprise Voice)'
    try {
        Import-Module MicrosoftTeams
        Connect-MicrosoftTeams | Out-Null
        Get-CsOnlineUser -Filter 'EnterpriseVoiceEnabled -eq $true' -ResultSize ([int]::MaxValue) |
            ForEach-Object { $voiceUsers[([string]$_.UserPrincipalName).ToLowerInvariant()] = [string]$_.LineUri }
        Write-Host "    $($voiceUsers.Count) Benutzer mit Teams-Telefonie"
    }
    catch {
        Write-Warning "Teams-Telefonie konnte nicht geprueft werden: $($_.Exception.Message)"
    }
}

# --------------------------------------------------------------------------------------------
# 5. Auswertung pro Benutzer
# --------------------------------------------------------------------------------------------
Write-Step 'Werte Benutzer aus'
$cutoff = (Get-Date).AddDays(-$InactiveDays)

$result = foreach ($u in $users) {
    $upn = ([string]$u.UserPrincipalName).ToLowerInvariant()

    # --- Lizenzen ---------------------------------------------------------------------------
    $partNumbers = @(foreach ($l in $u.AssignedLicenses) {
        $sku = $SkuById[[string]$l.SkuId]
        if ($sku) { $sku.SkuPartNumber } else { [string]$l.SkuId }
    })
    $has = @{}
    foreach ($cat in $SkuPattern.Keys) {
        $has[$cat] = [bool]($partNumbers | Where-Object { $_ -match $SkuPattern[$cat] })
    }

    # Zuweisung direkt oder ueber Gruppe?
    $viaGroup = @($u.LicenseAssignmentStates | Where-Object { $_.AssignedByGroup }).Count
    $direct   = @($u.LicenseAssignmentStates | Where-Object { -not $_.AssignedByGroup }).Count
    $assignment = if (-not $partNumbers) { '' }
                  elseif ($viaGroup -and $direct) { 'Gemischt' }
                  elseif ($viaGroup) { 'Gruppe' }
                  else { 'Direkt' }
    $assignErrors = @($u.LicenseAssignmentStates | Where-Object { $_.State -eq 'Error' } |
        ForEach-Object { $_.Error }) -join ', '

    # --- Anmeldung --------------------------------------------------------------------------
    $lastInteractive    = ConvertTo-Date $u.SignInActivity.LastSignInDateTime
    $lastNonInteractive = ConvertTo-Date $u.SignInActivity.LastNonInteractiveSignInDateTime
    $lastSignIn = @($lastInteractive, $lastNonInteractive) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
    $daysSinceSignIn = Get-DaysSince $lastSignIn

    # --- Nutzung ----------------------------------------------------------------------------
    $a = $activeUsers[$upn]
    $exchangeLast   = ConvertTo-Date $a.'Exchange Last Activity Date'
    $oneDriveLast   = ConvertTo-Date $a.'OneDrive Last Activity Date'
    $sharePointLast = ConvertTo-Date $a.'SharePoint Last Activity Date'
    $teamsLast      = ConvertTo-Date $a.'Teams Last Activity Date'

    $desktopActivations = [int]$activations[$upn]

    $t = $teamsUsage[$upn]
    $teamsCalls    = if ($t) { [int]("0" + $t.'Call Count') } else { 0 }
    $teamsMeetings = if ($t) { [int]("0" + $t.'Meeting Count') } else { 0 }

    $usesExchange   = $exchangeLast   -and $exchangeLast   -ge $cutoff
    $usesOneDrive   = $oneDriveLast   -and $oneDriveLast   -ge $cutoff
    $usesSharePoint = $sharePointLast -and $sharePointLast -ge $cutoff
    $usesTeams      = $teamsLast      -and $teamsLast      -ge $cutoff
    $usesM365       = $usesExchange -or $usesOneDrive -or $usesSharePoint -or $usesTeams
    $hasVoice       = $voiceUsers.ContainsKey($upn)

    # --- Empfehlung -------------------------------------------------------------------------
    $empfehlung = ''
    $begruendung = New-Object System.Collections.Generic.List[string]

    if ($partNumbers.Count -gt 0) {
        # Doppellizenzen: E5 enthaelt bereits Entra P1/P2, E3, EMS und Power BI Pro
        if ($has.E5) {
            $dupes = @($partNumbers | Where-Object {
                $_ -match $SkuPattern.EntraP2 -or $_ -match $SkuPattern.EntraP1 -or
                $_ -match $SkuPattern.E3 -or $_ -match $SkuPattern.EMS -or $_ -match $SkuPattern.PowerBI
            } | ForEach-Object { Get-SkuName $_ })
            if ($dupes) { $begruendung.Add("Doppellizenz (in E5 enthalten): $($dupes -join ', ')") }
        }

        if (-not $u.AccountEnabled) {
            $empfehlung = 'Lizenz entfernen'
            $begruendung.Add('Konto ist deaktiviert')
        }
        elseif ($null -eq $lastSignIn -and (ConvertTo-Date $u.CreatedDateTime) -and (ConvertTo-Date $u.CreatedDateTime) -lt $cutoff) {
            $empfehlung = 'Lizenz entfernen / pruefen'
            $begruendung.Add('Noch nie angemeldet')
        }
        elseif ($null -ne $daysSinceSignIn -and $daysSinceSignIn -gt $InactiveDays) {
            $empfehlung = 'Lizenz entfernen / pruefen'
            $begruendung.Add("Seit $daysSinceSignIn Tagen nicht angemeldet")
        }
        elseif ($has.E5) {
            if ($hasVoice) {
                $empfehlung = 'E5 behalten'
                $begruendung.Add('Nutzt Teams-Telefonie')
            }
            elseif (-not $usesM365 -and $desktopActivations -eq 0) {
                $empfehlung = 'Kandidat E5 -> Entra ID P2'
                $begruendung.Add("Keine Nutzung von Mail/Teams/OneDrive/SharePoint und keine Office-Desktop-Aktivierung (nur Anmeldung, z. B. an Dritttools)")
            }
            elseif ($desktopActivations -eq 0 -and -not $usesOneDrive) {
                $empfehlung = 'Pruefen: E5 -> kleinere Lizenz (F3/E3 + P2)'
                $begruendung.Add('Nur Web-/Mail-/Teams-Nutzung, keine Office-Desktop-Apps')
            }
            else {
                $empfehlung = 'E5 behalten'
                if ($desktopActivations -gt 0) { $begruendung.Add("Office-Desktop auf $desktopActivations Geraet(en) aktiviert") }
                if ($usesExchange) { $begruendung.Add('Nutzt Mailbox') }
                if ($usesTeams) { $begruendung.Add('Nutzt Teams') }
            }
        }
        else {
            $empfehlung = 'OK'
        }
    }

    [pscustomobject]@{
        Name                    = $u.DisplayName
        UPN                     = $u.UserPrincipalName
        Aktiviert               = $u.AccountEnabled
        Typ                     = $u.UserType
        EmployeeType            = $u.EmployeeType
        Firma                   = $u.CompanyName
        Abteilung               = $u.Department
        Funktion                = $u.JobTitle
        AD_OU                   = Get-OuFromDn $u.OnPremisesDistinguishedName
        AD_Synchronisiert       = [bool]$u.OnPremisesSyncEnabled
        Erstellt                = $u.CreatedDateTime
        Lizenzen                = ($partNumbers | ForEach-Object { Get-SkuName $_ }) -join ' | '
        Hat_E5                  = $has.E5
        Hat_E3                  = $has.E3
        Hat_F3                  = $has.F3
        Hat_EntraP2             = $has.EntraP2
        Hat_EntraP1             = $has.EntraP1
        Hat_LTSC                = $has.LTSC
        Zuweisung               = $assignment
        Zuweisungsfehler        = $assignErrors
        Letzte_Anmeldung        = $lastSignIn
        Tage_seit_Anmeldung     = $daysSinceSignIn
        Exchange_Letzte_Nutzung = $exchangeLast
        Teams_Letzte_Nutzung    = $teamsLast
        OneDrive_Letzte_Nutzung = $oneDriveLast
        SharePoint_Letzte_Nutzung = $sharePointLast
        Office_Desktop_Aktivierungen = $desktopActivations
        Teams_Anrufe            = $teamsCalls
        Teams_Meetings          = $teamsMeetings
        Teams_Telefonie         = $hasVoice
        Empfehlung              = $empfehlung
        Begruendung             = $begruendung -join ' / '
    }
}

# --------------------------------------------------------------------------------------------
# 6. Export
# --------------------------------------------------------------------------------------------
Write-Step 'Exportiere'
$licensed   = @($result | Where-Object { $_.Lizenzen })
$candidates = @($licensed | Where-Object { $_.Hat_E5 -and ($_.Empfehlung -ne 'E5 behalten' -or $_.Begruendung -like '*Doppellizenz*') })

Export-Report $skuOverview '01_SKU-Uebersicht'
Export-Report ($result | Sort-Object Empfehlung, Name) '02_Benutzer-Lizenzen'
Export-Report ($candidates | Sort-Object Empfehlung, Name) '03_E5-Kandidaten'

if (Get-Module -ListAvailable -Name ImportExcel) {
    $xlsx = Join-Path $OutputPath 'LizenzReport.xlsx'
    $skuOverview | Export-Excel -Path $xlsx -WorksheetName 'SKU-Uebersicht' -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet
    $result      | Export-Excel -Path $xlsx -WorksheetName 'Benutzer' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -ClearSheet
    $candidates  | Export-Excel -Path $xlsx -WorksheetName 'E5-Kandidaten' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -ClearSheet
    $result | Where-Object Lizenzen |
        Export-Excel -Path $xlsx -WorksheetName 'Pivot' -IncludePivotTable -PivotRows 'EmployeeType' -PivotColumns 'Empfehlung' -PivotData @{ UPN = 'Count' } -ClearSheet
    Write-Host "    $xlsx"
}

# --------------------------------------------------------------------------------------------
# 7. Zusammenfassung
# --------------------------------------------------------------------------------------------
Write-Step 'Zusammenfassung'
$skuOverview | Sort-Object Zugewiesen -Descending | Format-Table Produkt, Gekauft, Zugewiesen, Frei -AutoSize

Write-Host 'E5-Benutzer nach Empfehlung:' -ForegroundColor Yellow
$licensed | Where-Object Hat_E5 | Group-Object Empfehlung | Sort-Object Count -Descending |
    Format-Table @{ n = 'Empfehlung'; e = { $_.Name } }, Count -AutoSize

Write-Host 'E5-Benutzer nach EmployeeType / Firma (zum Abgleich mit den Personas):' -ForegroundColor Yellow
$licensed | Where-Object Hat_E5 | Group-Object EmployeeType, Firma | Sort-Object Count -Descending |
    Select-Object -First 20 | Format-Table @{ n = 'EmployeeType, Firma'; e = { $_.Name } }, Count -AutoSize

$dupeCount = @($licensed | Where-Object { $_.Begruendung -like '*Doppellizenz*' }).Count
Write-Host "Benutzer mit Doppellizenzen: $dupeCount"
Write-Host "Fertig. Berichte in: $OutputPath" -ForegroundColor Green
