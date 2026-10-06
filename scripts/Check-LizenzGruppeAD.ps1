<#
 Check-LizenzGruppeAD.ps1
 ------------------------
 Prueft die Mitglieder einer lokalen AD-Lizenzgruppe mit Get-ADUser und sagt pro Konto,
 ob die Lizenz bleiben soll oder entzogen werden kann.

 Laeuft in der PowerShell ISE auf einem Server mit AD-Modul (RSAT). Nur lesend.
 Es wird NICHTS geaendert, solange der Block ganz unten auskommentiert bleibt.

 Ausgabe: ein Ordner Lizenzcheck_<Datum> mit
   Mitglieder.csv      alle Mitglieder mit Status und Massnahme
   Kandidaten.csv      nur die Konten, bei denen etwas zu tun ist
   Zusammenfassung.txt Zahlen pro Status, EmployeeType und OU

 Wichtig zum Verstaendnis:
   LastLogonDate im AD ist die letzte Anmeldung an der DOMAENE (Windows-Login am Kantonslaptop,
   VDI, Netzlaufwerk). Wer sich nur in der Cloud anmeldet (Handy, privater PC), erscheint hier nicht.
   LastLogonDate wird nur alle ~14 Tage repliziert, also nie auf den Tag genau nehmen.
#>

# =============================================================================================
# 1. EINSTELLUNGEN - hier anpassen
# =============================================================================================
$Gruppe         = 'LIC-M365-E5'   # Name der AD-Lizenzgruppe
$InaktivTage    = 90              # ab so vielen Tagen ohne Domaenen-Anmeldung gilt ein Konto als inaktiv
$EintrittTage   = 60              # deaktivierte Konten, die juenger sind und nie angemeldet: vorbereiteter Eintritt
$Ausgabe        = "C:\Temp\Lizenzcheck_$(Get-Date -Format 'yyyy-MM-dd')"
$Ausnahmen      = @()             # Konten, die nie entzogen werden, z. B. @('svc_scan','m.muster')  (sAMAccountName)

# =============================================================================================
# 2. VORBEREITUNG
# =============================================================================================
Import-Module ActiveDirectory
New-Item -ItemType Directory -Path $Ausgabe -Force | Out-Null
$Heute = Get-Date

# =============================================================================================
# 3. MITGLIEDER DER GRUPPE LESEN (auch verschachtelte Gruppen)
# =============================================================================================
Write-Host "Lese Mitglieder von $Gruppe ..." -ForegroundColor Cyan
$Mitglieder = Get-ADGroupMember -Identity $Gruppe -Recursive | Where-Object objectClass -eq 'user'
Write-Host "  $($Mitglieder.Count) Benutzer in der Gruppe"

# =============================================================================================
# 4. PRO BENUTZER: DATEN AUS DEM AD HOLEN UND BEWERTEN
# =============================================================================================
$Ergebnis = foreach ($m in $Mitglieder) {

    $u = Get-ADUser -Identity $m.distinguishedName -Properties Enabled, whenCreated, LastLogonDate, PasswordLastSet,
            AccountExpirationDate, employeeType, department, mail, userPrincipalName, extensionAttribute10, memberOf

    # --- Zahlen, die wir fuer die Bewertung brauchen ------------------------------------------
    $TageSeitLogin   = if ($u.LastLogonDate) { [int]($Heute - $u.LastLogonDate).TotalDays } else { $null }
    $TageSeitErstellt = [int]($Heute - $u.whenCreated).TotalDays
    $OU = ($u.DistinguishedName -split ',' | Where-Object { $_ -like 'OU=*' } | ForEach-Object { $_.Substring(3) }) -join '/'

    # --- Bewertung, von oben nach unten: die erste zutreffende Regel gilt --------------------
    if ($u.SamAccountName -in $Ausnahmen) {
        $Status = 'Ausnahmeliste';                     $Massnahme = 'Kein Entzug'
    }
    elseif (-not $u.Enabled -and $TageSeitErstellt -le $EintrittTage -and -not $u.LastLogonDate) {
        $Status = 'Vorbereiteter Eintritt';            $Massnahme = 'Kein Entzug (Lizenz erst ab erstem Arbeitstag)'
    }
    elseif (-not $u.Enabled) {
        $Status = 'Konto deaktiviert (Austritt)';      $Massnahme = 'Entziehen (sofort)'
    }
    elseif ($u.AccountExpirationDate -and $u.AccountExpirationDate -lt $Heute) {
        $Status = 'Konto abgelaufen';                  $Massnahme = 'Entziehen (sofort)'
    }
    elseif (-not $u.LastLogonDate -and $TageSeitErstellt -gt $InaktivTage) {
        $Status = 'Nie angemeldet (Konto > 90 Tage)';  $Massnahme = 'Mit Amt klaeren, dann entziehen'
    }
    elseif (-not $u.LastLogonDate) {
        $Status = 'Neu, noch nie angemeldet';          $Massnahme = 'Abwarten'
    }
    elseif ($TageSeitLogin -gt 365) {
        $Status = 'Inaktiv > 365 Tage';                $Massnahme = 'Entziehen (sofort)'
    }
    elseif ($TageSeitLogin -gt 180) {
        $Status = 'Inaktiv 181-365 Tage';              $Massnahme = 'Mit Amt klaeren, dann entziehen'
    }
    elseif ($TageSeitLogin -gt $InaktivTage) {
        $Status = "Inaktiv $($InaktivTage+1)-180 Tage"; $Massnahme = 'Mit Amt klaeren (Langzeitabwesenheit?)'
    }
    else {
        $Status = 'Aktiv';                             $Massnahme = 'Lizenz bleibt'
    }

    # --- eine Zeile pro Benutzer ---------------------------------------------------------------
    [pscustomobject]@{
        Name              = $u.Name
        Kuerzel           = $u.SamAccountName
        UPN               = $u.UserPrincipalName
        Mail              = $u.mail
        Aktiviert         = $u.Enabled
        EmployeeType      = $u.employeeType
        Abteilung         = $u.department
        OU                = $OU
        Erstellt          = $u.whenCreated.ToString('dd.MM.yyyy')
        LetzterLogin      = if ($u.LastLogonDate) { $u.LastLogonDate.ToString('dd.MM.yyyy') } else { '' }
        TageSeitLogin     = $TageSeitLogin
        PasswortGesetzt   = if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString('dd.MM.yyyy') } else { '' }
        KontoLaeuftAb     = if ($u.AccountExpirationDate) { $u.AccountExpirationDate.ToString('dd.MM.yyyy') } else { '' }
        Status            = $Status
        Massnahme         = $Massnahme
        DN                = $u.DistinguishedName
    }
}

# =============================================================================================
# 5. AUSGABE
# =============================================================================================
$Kandidaten = $Ergebnis | Where-Object { $_.Massnahme -like 'Entziehen*' -or $_.Massnahme -like 'Mit Amt*' }

$Ergebnis   | Sort-Object Status, Name | Export-Csv "$Ausgabe\Mitglieder.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
$Kandidaten | Sort-Object Massnahme, Status, Name | Export-Csv "$Ausgabe\Kandidaten.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8

$Text = @()
$Text += "Lizenzcheck $Gruppe - $($Heute.ToString('dd.MM.yyyy HH:mm'))"
$Text += "Mitglieder: $($Ergebnis.Count)   Kandidaten: $($Kandidaten.Count)   davon sofort: $(@($Kandidaten | Where-Object Massnahme -like 'Entziehen*').Count)"
$Text += ''
$Text += 'Nach Status:'
$Text += ($Ergebnis | Group-Object Status | Sort-Object Count -Descending | Select-Object @{n="Wert";e={$_.Name}}, @{n="Anzahl";e={$_.Count}} | Format-Table | Out-String -Width 120)
$Text += 'Kandidaten nach EmployeeType:'
$Text += ($Kandidaten | Group-Object EmployeeType | Sort-Object Count -Descending | Select-Object @{n="Wert";e={$_.Name}}, @{n="Anzahl";e={$_.Count}} | Format-Table | Out-String -Width 120)
$Text += 'Kandidaten nach OU (oberste Ebene):'
$Text += ($Kandidaten | Group-Object { ($_.OU -split '/')[0] } | Sort-Object Count -Descending | Select-Object @{n="Wert";e={$_.Name}}, @{n="Anzahl";e={$_.Count}} | Format-Table | Out-String -Width 120)
$Text += "Dateien in: $Ausgabe"

$Text | Out-File "$Ausgabe\Zusammenfassung.txt" -Encoding UTF8
$Text | Write-Host
Write-Host 'Fertig. Es wurde nichts geaendert.' -ForegroundColor Green

# =============================================================================================
# 6. ENTZUG (bewusst auskommentiert)
#    Erst ausfuehren, wenn Kandidaten.csv geprueft ist. -WhatIf zeigt nur, was passieren wuerde.
# =============================================================================================
# $Sofort = Import-Csv "$Ausgabe\Kandidaten.csv" -Delimiter ';' | Where-Object Massnahme -like 'Entziehen*'
# foreach ($k in $Sofort) {
#     Remove-ADGroupMember -Identity $Gruppe -Members $k.Kuerzel -Confirm:$false -WhatIf
# }
