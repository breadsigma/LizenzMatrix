$Group          = 'LIC-M365-E5'
$InactiveDays   = 90
$NewAccountDays = 60
$OutputPath     = "C:\Temp\LicenseCheck_$(Get-Date -Format 'yyyy-MM-dd')"
$Exceptions     = @()

Import-Module ActiveDirectory
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
$Today = Get-Date

Write-Host "Reading members of $Group ..." -ForegroundColor Cyan
$Members = Get-ADGroupMember -Identity $Group -Recursive | Where-Object objectClass -eq 'user'
Write-Host "  $($Members.Count) users in group"

$Result = foreach ($m in $Members) {

    $u = Get-ADUser -Identity $m.distinguishedName -Properties Enabled, whenCreated, LastLogonDate, PasswordLastSet,
            AccountExpirationDate, employeeType, department, mail, userPrincipalName

    $DaysSinceLogin   = if ($u.LastLogonDate) { [int]($Today - $u.LastLogonDate).TotalDays } else { $null }
    $DaysSinceCreated = [int]($Today - $u.whenCreated).TotalDays
    $OU = ($u.DistinguishedName -split ',' | Where-Object { $_ -like 'OU=*' } | ForEach-Object { $_.Substring(3) }) -join '/'

    if ($u.SamAccountName -in $Exceptions) {
        $Status = 'Exception list';                    $Action = 'Keep'
    }
    elseif (-not $u.Enabled -and $DaysSinceCreated -le $NewAccountDays -and -not $u.LastLogonDate) {
        $Status = 'Pre-created new hire';              $Action = 'Keep (license from first working day)'
    }
    elseif (-not $u.Enabled) {
        $Status = 'Account disabled (leaver)';         $Action = 'Remove (immediately)'
    }
    elseif ($u.AccountExpirationDate -and $u.AccountExpirationDate -lt $Today) {
        $Status = 'Account expired';                   $Action = 'Remove (immediately)'
    }
    elseif (-not $u.LastLogonDate -and $DaysSinceCreated -gt $InactiveDays) {
        $Status = "Never logged in (account > $InactiveDays days)"; $Action = 'Check with department, then remove'
    }
    elseif (-not $u.LastLogonDate) {
        $Status = 'New, never logged in';              $Action = 'Wait'
    }
    elseif ($DaysSinceLogin -gt 365) {
        $Status = 'Inactive > 365 days';               $Action = 'Remove (immediately)'
    }
    elseif ($DaysSinceLogin -gt 180) {
        $Status = 'Inactive 181-365 days';             $Action = 'Check with department, then remove'
    }
    elseif ($DaysSinceLogin -gt $InactiveDays) {
        $Status = "Inactive $($InactiveDays + 1)-180 days"; $Action = 'Check with department (long-term absence?)'
    }
    else {
        $Status = 'Active';                            $Action = 'Keep'
    }

    [pscustomobject]@{
        SamAccountName = $u.SamAccountName
        UPN            = $u.UserPrincipalName
        EmployeeType   = $u.employeeType
        OU             = $OU
        Created        = $u.whenCreated.ToString('dd.MM.yyyy')
        LastLogin      = if ($u.LastLogonDate) { $u.LastLogonDate.ToString('dd.MM.yyyy') } else { 'Never' }
        LastSignInDays = if ($null -ne $DaysSinceLogin) { $DaysSinceLogin } else { 'Never' }
        LastPWSet      = if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString('dd.MM.yyyy') } else { 'Never' }
        AccountExpired = if ($u.AccountExpirationDate) { $u.AccountExpirationDate.ToString('dd.MM.yyyy') } else { 'Never' }
        Status         = $Status
        Action         = $Action
    }
}

$Candidates = $Result | Where-Object { $_.Action -like 'Remove*' -or $_.Action -like 'Check*' }
$Columns = 'SamAccountName', 'UPN', 'EmployeeType', 'OU', 'Created', 'LastLogin', 'LastSignInDays', 'LastPWSet', 'AccountExpired'

$Result     | Sort-Object SamAccountName | Select-Object $Columns |
    Export-Csv "$OutputPath\Members.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
$Candidates | Sort-Object Action, Status, SamAccountName | Select-Object ($Columns + 'Status', 'Action') |
    Export-Csv "$OutputPath\Candidates.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8

$Text = @()
$Text += "License check $Group - $($Today.ToString('dd.MM.yyyy HH:mm'))"
$Text += "Members: $($Result.Count)   Candidates: $($Candidates.Count)   Immediate: $(@($Candidates | Where-Object Action -like 'Remove*').Count)"
$Text += ''
$Text += 'By status:'
$Text += ($Result | Group-Object Status | Sort-Object Count -Descending | Select-Object @{n='Status';e={$_.Name}}, @{n='Count';e={$_.Count}} | Format-Table | Out-String -Width 120)
$Text += 'Candidates by EmployeeType:'
$Text += ($Candidates | Group-Object EmployeeType | Sort-Object Count -Descending | Select-Object @{n='EmployeeType';e={$_.Name}}, @{n='Count';e={$_.Count}} | Format-Table | Out-String -Width 120)
$Text += 'Candidates by OU (top level):'
$Text += ($Candidates | Group-Object { ($_.OU -split '/')[0] } | Sort-Object Count -Descending | Select-Object @{n='OU';e={$_.Name}}, @{n='Count';e={$_.Count}} | Format-Table | Out-String -Width 120)
$Text += "Output: $OutputPath"

$Text | Out-File "$OutputPath\Summary.txt" -Encoding UTF8
$Text | Write-Host
Write-Host 'Done. Nothing was changed.' -ForegroundColor Green

# $Immediate = Import-Csv "$OutputPath\Candidates.csv" -Delimiter ';' | Where-Object Action -like 'Remove*'
# foreach ($c in $Immediate) {
#     Remove-ADGroupMember -Identity $Group -Members $c.SamAccountName -Confirm:$false -WhatIf
# }
