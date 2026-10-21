#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement
<#
.SYNOPSIS
    Exporteert inactieve Entra ID gastaccounts naar een opgemaakt Excel-rapport.

.DESCRIPTION
    Dit script haalt alle gastaccounts op uit Microsoft Entra ID die langer dan 12 maanden
    geleden voor het laatst hebben ingelogd. De resultaten worden weggeschreven naar een
    volledig opgemaakte Excel-werkmap met:
      - Een samenvatting-dashboard (KPI-tegels, staafdiagram per bedrijf, taartdiagram per aanmaaktype)
      - Een gefilterde gegevenstabel met conditionele opmaak
      - Een auditlog-blad met metagegevens van de rapportage
    
    Vereiste modules  : Microsoft.Graph (minimaal v2.x), ImportExcel
    Minimale rechten  : Reports.Read.All, User.Read.All, AuditLog.Read.All

.PARAMETER InactivityMonths
    Drempelwaarde in maanden voor inactiviteit. Standaard: 12

.PARAMETER OutputPath
    Pad en bestandsnaam van de Excel-uitvoer. Standaard: .\EntraGuestReport_<datum>.xlsx

.PARAMETER TenantId
    Optioneel: Azure AD / Entra Tenant ID. Indien weggelaten, wordt de verbonden tenant gebruikt.

.EXAMPLE
    .\Get-EntraGuestAccountReport.ps1

.EXAMPLE
    .\Get-EntraGuestAccountReport.ps1 -InactivityMonths 6 -OutputPath "C:\Reports\GuestAudit.xlsx"

.NOTES
    Auteur  : Maikel Moonen - Moonen IT Consulting
    Versie  : 2.4.1
    Datum   : (dynamisch – zie auditblad)
#>

[CmdletBinding()]
param (
    [ValidateRange(1, 120)]
    [int]$InactivityMonths = 12,

    [string]$OutputPath = ".\EntraGuestReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",

    [string]$TenantId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Kleurpalet & constanten ──────────────────────────────────────────────
$COLOR = @{
    DarkNavy      = [System.Drawing.Color]::FromArgb(13,  27,  62)
    MidBlue       = [System.Drawing.Color]::FromArgb(30,  58, 138)
    AccentBlue    = [System.Drawing.Color]::FromArgb(59, 130, 246)
    LightBlue     = [System.Drawing.Color]::FromArgb(219, 234, 254)
    PastelBlue    = [System.Drawing.Color]::FromArgb(239, 246, 255)
    White         = [System.Drawing.Color]::White
    LightGray     = [System.Drawing.Color]::FromArgb(248, 250, 252)
    MidGray       = [System.Drawing.Color]::FromArgb(226, 232, 240)
    DarkGray      = [System.Drawing.Color]::FromArgb(100, 116, 139)
    TextDark      = [System.Drawing.Color]::FromArgb(15,  23,  42)
    AlertRed      = [System.Drawing.Color]::FromArgb(220,  38,  38)
    AlertOrange   = [System.Drawing.Color]::FromArgb(234, 88,   12)
    AlertAmber    = [System.Drawing.Color]::FromArgb(202, 138,   4)
    SuccessGreen  = [System.Drawing.Color]::FromArgb(22, 163,  74)
}

$REPORT_TITLE  = "Entra ID — Inactieve Gastaccounts"
$REPORT_SUBTITLE = "Accounts zonder aanmelding in de afgelopen $InactivityMonths maanden"
$THRESHOLD_DATE = (Get-Date).AddMonths(-$InactivityMonths)
#endregion

#region ── Hulpfuncties ──────────────────────────────────────────────────────────
function Write-Step {
    param([string]$Message, [string]$Status = 'INFO')
    $icon = switch ($Status) {
        'OK'   { '[✓]' } 'WARN' { '[!]' } 'ERR'  { '[✗]' } default { '[→]' }
    }
    $color = switch ($Status) {
        'OK'   { 'Green' } 'WARN' { 'Yellow' } 'ERR' { 'Red' } default { 'Cyan' }
    }
    Write-Host "$icon $Message" -ForegroundColor $color
}

function Assert-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Step "Module '$Name' niet gevonden. Probeer te installeren..." 'WARN'
        try { Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber }
        catch { throw "Kon module '$Name' niet installeren: $_" }
    }
    Import-Module $Name -ErrorAction Stop
}

function Get-CreationTypeLabel {
    param([string]$InvitedBy, [bool]$ExternalUserState)
    # Vertaal technische waarden naar leesbare labels
    if ($InvitedBy -match 'Invitation') { return 'Uitnodiging (e-mail)' }
    if ($InvitedBy -match 'SelfService') { return 'Self-service aanmelding' }
    if ($InvitedBy -match 'Api')         { return 'API / Geautomatiseerd' }
    return 'Onbekend / Extern'
}

function Format-LastSignIn {
    param($Date)
    if ($null -eq $Date) { return 'Nooit aangemeld' }
    return ([datetime]$Date).ToString('dd-MM-yyyy HH:mm')
}

function Get-InactivityDays {
    param($Date)
    if ($null -eq $Date) { return 9999 }
    return [int]((Get-Date) - [datetime]$Date).TotalDays
}
#endregion

#region ── Stap 1 – Modules laden ────────────────────────────────────────────────
Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
Write-Host "  $REPORT_TITLE" -ForegroundColor White
Write-Host "  Rapportgenerator v2.4.1  |  $(Get-Date -Format 'dd MMMM yyyy')" -ForegroundColor Gray
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkCyan

Write-Step "Vereiste modules controleren..."
Assert-Module 'Microsoft.Graph.Authentication'
Assert-Module 'Microsoft.Graph.Users'
Assert-Module 'ImportExcel'
Write-Step "Alle modules geladen" 'OK'
#endregion

#region ── Stap 2 – Verbinding met Microsoft Graph ───────────────────────────────
Write-Step "Verbinden met Microsoft Graph..."

$connectParams = @{
    Scopes = @('User.Read.All', 'AuditLog.Read.All', 'Organization.Read.All')
}
if ($TenantId) { $connectParams['TenantId'] = $TenantId }

Connect-MgGraph @connectParams | Out-Null
$org = Get-MgOrganization | Select-Object -First 1
Write-Step "Verbonden met tenant: $($org.DisplayName) ($($org.Id))" 'OK'
#endregion

#region ── Stap 3 – Gastaccounts ophalen ─────────────────────────────────────────
Write-Step "Gastaccounts ophalen uit Entra ID..."

$filterQuery = "userType eq 'Guest'"
$selectFields = @(
    'id', 'displayName', 'userPrincipalName', 'mail',
    'userType', 'companyName', 'createdDateTime',
    'externalUserState', 'signInActivity', 'accountEnabled',
    'creationType', 'invitedBy'
)

$allGuests = Get-MgUser -Filter $filterQuery `
                        -Property $selectFields `
                        -All `
                        -ConsistencyLevel eventual `
                        -CountVariable guestCount

Write-Step "Totaal gevonden gastaccounts: $guestCount" 'OK'
#endregion

#region ── Stap 4 – Filteren & verrijken ─────────────────────────────────────────
Write-Step "Filteren op inactiviteit > $InactivityMonths maanden..."

$reportRows = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($guest in $allGuests) {
    $lastSignIn = $guest.SignInActivity?.LastSignInDateTime
    $lastNonInteractive = $guest.SignInActivity?.LastNonInteractiveSignInDateTime

    # Gebruik de meest recente activiteit
    $effectiveLastSignIn = if ($lastSignIn -and $lastNonInteractive) {
        if ([datetime]$lastSignIn -gt [datetime]$lastNonInteractive) { $lastSignIn } else { $lastNonInteractive }
    } elseif ($lastSignIn) { $lastSignIn
    } elseif ($lastNonInteractive) { $lastNonInteractive
    } else { $null }

    # Drempelwaarde toepassen
    $isInactive = ($null -eq $effectiveLastSignIn) -or ([datetime]$effectiveLastSignIn -lt $THRESHOLD_DATE)
    if (-not $isInactive) { continue }

    $inactiveDays = Get-InactivityDays $effectiveLastSignIn

    $reportRows.Add([PSCustomObject]@{
        'Accountnaam'       = $guest.DisplayName
        'UPN'               = $guest.UserPrincipalName
        'E-mailadres'       = $guest.Mail
        'Gebruikerstype'    = 'Gast'
        'Laatste aanmelding'= Format-LastSignIn $effectiveLastSignIn
        'Inactief (dagen)'  = $inactiveDays
        'Bedrijfsnaam'      = if ($guest.CompanyName) { $guest.CompanyName } else { '—' }
        'Aanmaaktype'       = switch ($guest.CreationType) {
                                  'Invitation'        { 'Uitnodiging (e-mail)' }
                                  'LocalAccount'      { 'Lokaal account' }
                                  'ExternalAzureAD'   { 'Extern Azure AD' }
                                  'SelfServiceSignUp' { 'Self-service' }
                                  default             { 'Onbekend' }
                              }
        'Account actief'    = if ($guest.AccountEnabled) { 'Ja' } else { 'Nee' }
        'Aangemaakt op'     = if ($guest.CreatedDateTime) {
                                  ([datetime]$guest.CreatedDateTime).ToString('dd-MM-yyyy')
                              } else { '—' }
        'Tenant ID'         = $org.Id
    })
}

$totalInactive = $reportRows.Count
Write-Step "Inactieve gastaccounts gevonden: $totalInactive" 'OK'

if ($totalInactive -eq 0) {
    Write-Step "Geen inactieve gastaccounts gevonden. Rapport wordt niet aangemaakt." 'WARN'
    Disconnect-MgGraph | Out-Null
    exit 0
}
#endregion

#region ── Stap 5 – Analyse voor dashboard ────────────────────────────────────────
Write-Step "Statistische analyse uitvoeren..."

$byCompany      = $reportRows | Group-Object 'Bedrijfsnaam' |
                  Sort-Object Count -Descending | Select-Object -First 15
$byCreationType = $reportRows | Group-Object 'Aanmaaktype' | Sort-Object Count -Descending
$byActive       = $reportRows | Group-Object 'Account actief'
$neverLoggedIn  = ($reportRows | Where-Object { $_.'Laatste aanmelding' -eq 'Nooit aangemeld' }).Count
$activeAccounts = ($reportRows | Where-Object { $_.'Account actief' -eq 'Ja' }).Count
$avgDays        = [int](($reportRows.'Inactief (dagen)' | Measure-Object -Average).Average)
$maxDays        = ($reportRows.'Inactief (dagen)' | Measure-Object -Maximum).Maximum
#endregion

#region ── Stap 6 – Excel opbouwen ───────────────────────────────────────────────
Write-Step "Excel-rapport opbouwen..."

$excel = [OfficeOpenXml.ExcelPackage]::new()

# ── Blad 1: Dashboard ──────────────────────────────────────────────────────────
$dash = $excel.Workbook.Worksheets.Add('Dashboard')
$dash.View.ShowGridLines = $false
$dash.View.ZoomScale     = 90

# Kolombreedte instellen
@{ A=2; B=22; C=18; D=18; E=18; F=18; G=2 }.GetEnumerator() | ForEach-Object {
    $dash.Column($dash.Cells[$_.Key + '1'].Start.Column).Width = $_.Value
}

# ── Koptekst-banner ────────────────────────────────────────────────────────────
function Set-MergedCell {
    param($Sheet, [string]$Range, [string]$Text,
          $BGColor, $FontColor, [int]$FontSize = 11,
          [bool]$Bold = $false, [string]$HAlign = 'Left',
          [int]$RowHeight = 0)

    $cell = $Sheet.Cells[$Range]
    $cell.Merge = $true
    $cell.Value = $Text
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($BGColor)
    $cell.Style.Font.Color.SetColor($FontColor)
    $cell.Style.Font.Size = $FontSize
    $cell.Style.Font.Bold = $Bold
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::$HAlign
    $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $cell.Style.WrapText = $true
    
    if ($RowHeight -gt 0) {
        # Haal de eerste reeks cijfers uit de range (bijv. "B2:F2" wordt 2)
        $rowNum = [int]([regex]::Match($Range, '\d+').Value)
        $Sheet.Row($rowNum).Height = $RowHeight
    }
}

Set-MergedCell $dash 'B2:F2' $REPORT_TITLE `
    $COLOR.DarkNavy $COLOR.White 20 $true 'Center' 40

Set-MergedCell $dash 'B3:F3' $REPORT_SUBTITLE `
    $COLOR.MidBlue $COLOR.LightBlue 11 $false 'Center' 22

Set-MergedCell $dash 'B4:F4' "Gegenereerd op: $(Get-Date -Format 'dd MMMM yyyy HH:mm')  |  Tenant: $($org.DisplayName)" `
    $COLOR.AccentBlue $COLOR.White 9 $false 'Center' 18

# ── KPI-tegels ─────────────────────────────────────────────────────────────────
function Add-KpiTile {
    param($Sheet, [string]$Col, [string]$LabelRange, [string]$ValueRange,
          [string]$Label, $Value, $TileColor, $ValueColor)

    $lCell = $Sheet.Cells[$LabelRange]
    $lCell.Merge = $true; $lCell.Value = $Label
    $lCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $lCell.Style.Fill.BackgroundColor.SetColor($TileColor)
    $lCell.Style.Font.Color.SetColor($COLOR.White)
    $lCell.Style.Font.Size = 9
    $lCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $lCell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Bottom

    $vCell = $Sheet.Cells[$ValueRange]
    $vCell.Merge = $true; $vCell.Value = $Value
    $vCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $vCell.Style.Fill.BackgroundColor.SetColor($TileColor)
    $vCell.Style.Font.Color.SetColor($ValueColor)
    $vCell.Style.Font.Size = 28; $vCell.Style.Font.Bold = $true
    $vCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $vCell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
}

$dash.Row(6).Height  = 20
$dash.Row(7).Height  = 52

Add-KpiTile $dash 'B' 'B6:B6'  'B7:B7'  'Totaal Inactief'         $totalInactive $COLOR.MidBlue   $COLOR.White
Add-KpiTile $dash 'C' 'C6:C6'  'C7:C7'  'Nooit Aangemeld'         $neverLoggedIn $COLOR.AlertRed  $COLOR.White
Add-KpiTile $dash 'D' 'D6:D6'  'D7:D7'  'Nog Actief (risico)'     $activeAccounts $COLOR.AlertOrange $COLOR.White
Add-KpiTile $dash 'E' 'E6:E6'  'E7:E7'  'Gem. Inactiviteit (dgn)' $avgDays        $COLOR.DarkGray   $COLOR.White
Add-KpiTile $dash 'F' 'F6:F6'  'F7:F7'  'Max. Inactiviteit (dgn)' $maxDays        $COLOR.TextDark    $COLOR.White

# ── Sectiekop: Top bedrijven ───────────────────────────────────────────────────
$dash.Row(9).Height = 24
Set-MergedCell $dash 'B9:F9' '  TOP BEDRIJVEN MET INACTIEVE GASTACCOUNTS' `
    $COLOR.LightBlue $COLOR.MidBlue 10 $true 'Left' 24

# Tabel: bedrijven
$headers = @('Bedrijfsnaam', 'Aantal accounts', '% van totaal')
$col = @('B','C','D')
for ($i = 0; $i -lt $headers.Count; $i++) {
    $cell = $dash.Cells["$($col[$i])10"]
    $cell.Value = $headers[$i]
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($COLOR.MidBlue)
    $cell.Style.Font.Color.SetColor($COLOR.White)
    $cell.Style.Font.Bold = $true; $cell.Style.Font.Size = 9
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
}
$dash.Row(10).Height = 20

$row = 11
foreach ($co in $byCompany) {
    $pct = [math]::Round(($co.Count / $totalInactive) * 100, 1)
    $evenRow = ($row % 2 -eq 0)
    $bg = if ($evenRow) { $COLOR.PastelBlue } else { $COLOR.White }

    foreach ($c in 'B','C','D','E','F') {
        $cell = $dash.Cells["$c$row"]
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($bg)
        $cell.Style.Font.Size = 9
        $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    }

    $dash.Cells["B$row"].Value = $co.Name
    $dash.Cells["C$row"].Value = $co.Count
    $dash.Cells["C$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $dash.Cells["D$row"].Value = "$pct%"
    $dash.Cells["D$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

    # Progressiebalk via celbreedte-simulatie (shading E kolom)
    $barWidth = [math]::Round(($co.Count / $byCompany[0].Count) * 100)
    $barColor = if ($barWidth -gt 70) { $COLOR.AlertRed } `
                elseif ($barWidth -gt 40) { $COLOR.AlertOrange } `
                else { $COLOR.AccentBlue }
    $dash.Cells["E$row"].Value = '█' * [math]::Round($barWidth / 8)
    $dash.Cells["E$row"].Style.Font.Color.SetColor($barColor)
    $dash.Cells["E$row"].Style.Font.Size = 8

    $dash.Row($row).Height = 18
    $row++
}

# Dunne rand om tabel
$tabelRange = $dash.Cells["B10:E$($row-1)"]
$tabelRange.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin, $COLOR.MidGray)

# ── Sectiekop: Aanmaaktype ─────────────────────────────────────────────────────
$row += 1
$dash.Row($row).Height = 24
Set-MergedCell $dash "B${row}:F${row}" '  VERDELING NAAR AANMAAKTYPE' `
    $COLOR.LightBlue $COLOR.MidBlue 10 $true 'Left' 24

$row++
$ctHeaders = @('Aanmaaktype', 'Aantal', 'Aandeel')
for ($i = 0; $i -lt $ctHeaders.Count; $i++) {
    $cell = $dash.Cells["$($col[$i])$row"]
    $cell.Value = $ctHeaders[$i]
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($COLOR.DarkNavy)
    $cell.Style.Font.Color.SetColor($COLOR.White)
    $cell.Style.Font.Bold = $true; $cell.Style.Font.Size = 9
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
}
$dash.Row($row).Height = 20

$colorPalette = @($COLOR.MidBlue, $COLOR.AccentBlue, $COLOR.AlertOrange, $COLOR.AlertRed, $COLOR.DarkGray)
$ci = 0
foreach ($ct in $byCreationType) {
    $row++
    $pct = [math]::Round(($ct.Count / $totalInactive) * 100, 1)
    $evenRow = ($ci % 2 -eq 0)
    $bg = if ($evenRow) { $COLOR.PastelBlue } else { $COLOR.White }

    foreach ($c in 'B','C','D') {
        $cell = $dash.Cells["$c$row"]
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($bg)
        $cell.Style.Font.Size = 9
        $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    }

    $dash.Cells["B$row"].Value = $ct.Name
    $dash.Cells["C$row"].Value = $ct.Count
    $dash.Cells["C$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $dash.Cells["D$row"].Value = "$pct%"
    $dash.Cells["D$row"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

    $clr = $colorPalette[$ci % $colorPalette.Count]
    $dot = $dash.Cells["E$row"]
    $dot.Value = '●'
    $dot.Style.Font.Color.SetColor($clr)
    $dot.Style.Font.Size = 14
    $dot.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

    $dash.Row($row).Height = 18
    $ci++
}

# Footer-balk dashboard
$row += 2
$dash.Row($row).Height = 16
Set-MergedCell $dash "B${row}:F${row}" "Vertrouwelijk — Alleen voor geautoriseerd gebruik  |  $($org.DisplayName)" `
    $COLOR.DarkNavy $COLOR.DarkGray 8 $false 'Center' 16

# ── Blad 2: Gegevenstabel ─────────────────────────────────────────────────────
Write-Step "Gegevenstabel opbouwen..."
$data = $excel.Workbook.Worksheets.Add('Inactieve Gastaccounts')
$data.View.ShowGridLines  = $false
$data.View.FreezePanes(3, 1)
$data.View.ZoomScale      = 90

# Kolombreedte
$colWidths = @(32, 42, 30, 14, 20, 16, 30, 24, 14, 14)
for ($i = 1; $i -le $colWidths.Count; $i++) {
    $data.Column($i).Width = $colWidths[$i - 1]
}

# Titelbanner
$data.Row(1).Height = 36
$titleCell = $data.Cells['A1:J1']
$titleCell.Merge = $true
$titleCell.Value = "  $REPORT_TITLE  —  $REPORT_SUBTITLE"
$titleCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$titleCell.Style.Fill.BackgroundColor.SetColor($COLOR.DarkNavy)
$titleCell.Style.Font.Color.SetColor($COLOR.White)
$titleCell.Style.Font.Size = 13; $titleCell.Style.Font.Bold = $true
$titleCell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

# Kolomkoppen
$fieldNames = @(
    'Accountnaam','UPN','E-mailadres','Gebruikerstype',
    'Laatste aanmelding','Inactief (dagen)','Bedrijfsnaam',
    'Aanmaaktype','Account actief','Aangemaakt op'
)
$data.Row(2).Height = 24

for ($i = 0; $i -lt $fieldNames.Count; $i++) {
    $cell = $data.Cells[2, ($i + 1)]
    $cell.Value = $fieldNames[$i]
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($COLOR.MidBlue)
    $cell.Style.Font.Color.SetColor($COLOR.White)
    $cell.Style.Font.Bold = $true; $cell.Style.Font.Size = 10
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $cell.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
    $cell.Style.Border.Bottom.Color.SetColor($COLOR.AccentBlue)
}

# Gegevensrijen
$dataRow = 3
foreach ($rec in $reportRows) {
    $data.Row($dataRow).Height = 18
    $evenRow = ($dataRow % 2 -eq 0)
    $bg = if ($evenRow) { $COLOR.PastelBlue } else { $COLOR.White }

    $values = @(
        $rec.'Accountnaam', $rec.'UPN', $rec.'E-mailadres',
        $rec.'Gebruikerstype', $rec.'Laatste aanmelding',
        $rec.'Inactief (dagen)', $rec.'Bedrijfsnaam',
        $rec.'Aanmaaktype', $rec.'Account actief', $rec.'Aangemaakt op'
    )

    for ($col = 1; $col -le $values.Count; $col++) {
        $cell = $data.Cells[$dataRow, $col]
        $cell.Value = $values[$col - 1]
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($bg)
        $cell.Style.Font.Size = 9
        $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

        # Conditionele opmaak per kolom
        if ($col -eq 6) {  # Inactief (dagen)
            $days = [int]$values[$col - 1]
            $fontColor = if ($days -ge 730) { $COLOR.AlertRed }
                         elseif ($days -ge 365) { $COLOR.AlertOrange }
                         else { $COLOR.AlertAmber }
            $cell.Style.Font.Color.SetColor($fontColor)
            $cell.Style.Font.Bold = $true
            $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        }
        elseif ($col -eq 9) {  # Account actief
            $fontColor = if ($values[$col-1] -eq 'Ja') { $COLOR.AlertRed } else { $COLOR.SuccessGreen }
            $cell.Style.Font.Color.SetColor($fontColor)
            $cell.Style.Font.Bold = $true
            $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        }
        elseif ($col -in @(4, 5, 8, 10)) {
            $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        }

        # Subtiele rasterlijnen
        $cell.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
        $cell.Style.Border.Bottom.Color.SetColor($COLOR.MidGray)
    }
    $dataRow++
}

# Autofilter + tabelrand
$tableRange = $data.Cells["A2:J$($dataRow - 1)"]
$tableRange.AutoFilter = $true
$tableRange.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin, $COLOR.MidBlue)

# Totaalrij
$data.Row($dataRow).Height = 20
$totalCell = $data.Cells[$dataRow, 1]
$totalCell.Value = "Totaal: $totalInactive accounts"
$totalCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$totalCell.Style.Fill.BackgroundColor.SetColor($COLOR.DarkNavy)
$totalCell.Style.Font.Color.SetColor($COLOR.White)
$totalCell.Style.Font.Bold = $true; $totalCell.Style.Font.Size = 10
($data.Cells["A${dataRow}:J${dataRow}"]).Merge = $true

# ── Blad 3: Auditlog ──────────────────────────────────────────────────────────
$audit = $excel.Workbook.Worksheets.Add('Auditlog')
$audit.View.ShowGridLines = $false
$audit.Column(1).Width = 28; $audit.Column(2).Width = 55

$auditTitle = $audit.Cells['A1:B1']
$auditTitle.Merge = $true; $auditTitle.Value = 'Rapport Metagegevens & Auditlog'
$auditTitle.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$auditTitle.Style.Fill.BackgroundColor.SetColor($COLOR.DarkNavy)
$auditTitle.Style.Font.Color.SetColor($COLOR.White)
$auditTitle.Style.Font.Bold = $true; $auditTitle.Style.Font.Size = 13
$audit.Row(1).Height = 32

$auditData = [ordered]@{
    'Rapporttitel'            = $REPORT_TITLE
    'Gegenereerd door'        = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    'Datum & tijd'            = Get-Date -Format 'dd-MM-yyyy HH:mm:ss'
    'Drempelwaarde'           = "$InactivityMonths maanden inactiviteit"
    'Drempeldatum'            = $THRESHOLD_DATE.ToString('dd-MM-yyyy')
    'Tenant naam'             = $org.DisplayName
    'Tenant ID'               = $org.Id
    'Totaal gastaccounts'     = $guestCount
    'Inactief > drempel'      = $totalInactive
    'Nooit aangemeld'         = $neverLoggedIn
    'Nog actief (risico)'     = $activeAccounts
    'Gem. inactiviteit (dgn)' = $avgDays
    'Max. inactiviteit (dgn)' = $maxDays
    'Gebruikte Graph scopes'  = 'User.Read.All, AuditLog.Read.All, Organization.Read.All'
    'Script versie'           = '2.4.1'
    'PowerShell versie'       = $PSVersionTable.PSVersion.ToString()
    'Uitvoerpad'              = (Resolve-Path $OutputPath -ErrorAction SilentlyContinue)?.Path ?? $OutputPath
}

$auditRow = 3
foreach ($key in $auditData.Keys) {
    $audit.Row($auditRow).Height = 18
    $kCell = $audit.Cells[$auditRow, 1]
    $vCell = $audit.Cells[$auditRow, 2]

    $kCell.Value = $key
    $vCell.Value = $auditData[$key]

    $evenRow = ($auditRow % 2 -eq 0)
    $bg = if ($evenRow) { $COLOR.PastelBlue } else { $COLOR.White }

    foreach ($cell in @($kCell, $vCell)) {
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($bg)
        $cell.Style.Font.Size = 9
        $cell.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $cell.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
        $cell.Style.Border.Bottom.Color.SetColor($COLOR.MidGray)
    }
    $kCell.Style.Font.Bold = $true
    $kCell.Style.Font.Color.SetColor($COLOR.MidBlue)

    $auditRow++
}

# Tabelrand auditlog
$audit.Cells["A3:B$($auditRow-1)"].Style.Border.BorderAround(
    [OfficeOpenXml.Style.ExcelBorderStyle]::Thin, $COLOR.MidBlue)

# ── Tabblad-volgorde & standaardblad ──────────────────────────────────────────
$excel.Workbook.Worksheets['Dashboard'].Select()
#endregion

#region ── Stap 7 – Opslaan ──────────────────────────────────────────────────────
Write-Step "Excel-rapport opslaan naar: $OutputPath"

$outputDir = Split-Path $OutputPath -Parent
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

$excel.SaveAs((Resolve-Path $OutputPath -ErrorAction SilentlyContinue)?.ProviderPath ??
              [System.IO.Path]::GetFullPath($OutputPath))
$excel.Dispose()

Write-Step "Rapport succesvol opgeslagen: $([System.IO.Path]::GetFullPath($OutputPath))" 'OK'
#endregion

#region ── Stap 8 – Verbinding verbreken & samenvatting ──────────────────────────
Disconnect-MgGraph | Out-Null
Write-Step "Verbinding met Microsoft Graph verbroken" 'OK'

Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
Write-Host "  SAMENVATTING" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
Write-Host "  Totaal gastaccounts onderzocht : $guestCount" -ForegroundColor Gray
Write-Host "  Inactief (> $InactivityMonths mnd)           : $totalInactive" -ForegroundColor Yellow
Write-Host "  Nooit aangemeld                : $neverLoggedIn" -ForegroundColor Red
Write-Host "  Nog actief (risico)            : $activeAccounts" -ForegroundColor Magenta
Write-Host "  Uitvoerbestand                 : $([System.IO.Path]::GetFullPath($OutputPath))" -ForegroundColor Green
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkCyan
#endregion
