<#
.SYNOPSIS
    Genereert een Excel-dashboard met het Visio Plan 2 gebruik per gebruiker (Microsoft 365).

.DESCRIPTION
    BELANGRIJKE BEPERKING (lees dit eerst):
    Microsoft Graph biedt GEEN API-endpoint voor Visio (of Project) gebruiksrapportages.
    Workloads als Teams, Exchange, SharePoint en OneDrive hebben wel een Graph reports
    endpoint (bv. getTeamsUserActivityUserDetail), maar Visio niet. Dit is een bekende
    en bevestigde beperking van Microsoft Graph (zie o.a. Microsoft Q&A en de officiele
    Visio activity report documentatie).

    Daarom werkt dit script in twee stappen:
      1. HANDMATIG (eenmalig per rapportage): exporteer het Visio activiteitenrapport
         als CSV vanuit het Microsoft 365 admin center.
      2. AUTOMATISCH (dit script): combineert die CSV met de actuele Visio Plan 2
         licentietoewijzingen uit Entra ID (via Microsoft Graph) en bouwt daaruit een
         volledig opgemaakt Excel-dashboard met KPI's, kleurcodering, en grafieken.

    Door de licentietoewijzing (Graph) te combineren met de activiteitsdata (CSV) krijg
    je niet alleen "wie heeft Visio gebruikt", maar ook "wie heeft de licentie maar
    gebruikt 'm nooit" - vaak de meest waardevolle informatie voor licentie-optimalisatie.

    STAPPEN OM DE CSV TE VERKRIJGEN:
      1. Ga naar https://admin.microsoft.com
      2. Navigeer naar Rapporten (Reports) > Gebruik (Usage)
      3. Selecteer "Visio" in de lijst met rapporten
      4. Kies de gewenste periode (7/30/90/180 dagen) - 180 dagen geeft het meeste inzicht
      5. Klik op "Exporteren" om de CSV te downloaden
      6. Geef het pad van die CSV mee via -VisioActivityCsvPath, of laat het script
         zelf zoeken in je Downloads-map (bestandsnaam moet "visio" bevatten)

    VEREISTE GRAPH PERMISSIES (delegated): User.Read.All, Organization.Read.All
    VEREISTE ADMIN CENTER ROL: Reports Reader (of hoger) om het rapport te kunnen exporteren

.PARAMETER VisioActivityCsvPath
    Pad naar de geexporteerde Visio activity CSV. Indien niet opgegeven, wordt gezocht
    naar het meest recente bestand met "visio" in de naam in de Downloads-map.

.PARAMETER TenantId
    Tenant ID (GUID of domeinnaam) voor Connect-MgGraph. Optioneel.

.PARAMETER VisioSkuPartNumber
    SkuPartNumber van de Visio Plan 2 licentie in deze tenant. Standaard "VISIOCLIENT".
    Controleer dit zelf met Get-MgSubscribedSku als je twijfelt - licentienamen kunnen
    per tenant/historie afwijken.

.PARAMETER InactiveThresholdDays
    Aantal dagen zonder activiteit waarna een gebruiker als "Inactief" geldt. Standaard 30.

.PARAMETER Period
    Label van de rapportageperiode, alleen voor weergave op het dashboard (komt overeen
    met de periode die je bij export in de admin center hebt gekozen).

.PARAMETER OutputPath
    Volledig pad voor het output Excel-bestand. Standaard: .\Reports\Visio_Plan2_Gebruiksoverzicht_<timestamp>.xlsx

.PARAMETER OpenOutputFolder
    Open de map met het rapport na afronding.

.PARAMETER ShowExcelAfterCreation
    Open het Excel-bestand direct na het maken.

.EXAMPLE
    .\Invoke-VisioPlan2UsageReport.ps1 -VisioActivityCsvPath "C:\Downloads\VisioActivityUserDetail.csv" -Period D180

.EXAMPLE
    .\Invoke-VisioPlan2UsageReport.ps1 -InactiveThresholdDays 60 -OpenOutputFolder

.NOTES
    Auteur: Maikel Moonen - Moonen IT Consulting
    Vereist modules: Microsoft.Graph.Authentication, Microsoft.Graph.Users,
                      Microsoft.Graph.Identity.DirectoryManagement, ImportExcel
#>

[CmdletBinding()]
param(
    [Parameter(HelpMessage = 'Pad naar de geexporteerde Visio activity CSV. Indien leeg: zoeken in Downloads.')]
    [string]$VisioActivityCsvPath,

    [Parameter(HelpMessage = 'Tenant ID (GUID of domeinnaam) voor Connect-MgGraph.')]
    [string]$TenantId,

    [Parameter(HelpMessage = 'SkuPartNumber van de Visio Plan 2 licentie.')]
    [string]$VisioSkuPartNumber = 'VISIOCLIENT',

    [Parameter(HelpMessage = 'Aantal dagen zonder activiteit voordat een gebruiker als inactief geldt.')]
    [int]$InactiveThresholdDays = 30,

    [Parameter(HelpMessage = 'Label van de rapportageperiode, alleen voor weergave op het dashboard.')]
    [ValidateSet('D7', 'D30', 'D90', 'D180', 'Onbekend')]
    [string]$Period = 'D180',

    [Parameter(HelpMessage = 'Volledig pad voor het output Excel-bestand.')]
    [string]$OutputPath,

    [switch]$OpenOutputFolder,

    [switch]$ShowExcelAfterCreation
)

#region Logging

$script:LogFilePath = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')][string]$Level = 'Info'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'Warning' { Write-Warning $line }
        'Error'   { Write-Host $line -ForegroundColor Red }
        default   { Write-Host $line -ForegroundColor Cyan }
    }

    if ($script:LogFilePath) {
        try { Add-Content -Path $script:LogFilePath -Value $line -Encoding UTF8 } catch { }
    }
}

#endregion

#region Helpers

function Convert-ColumnNumberToLetter {
    param([int]$Number)
    $letter = ''
    while ($Number -gt 0) {
        $rem = ($Number - 1) % 26
        $letter = [char](65 + $rem) + $letter
        $Number = [int](($Number - $rem - 1) / 26)
    }
    return $letter
}

function Set-CellFill {
    # Kleurt een EPPlus cel solid met een hex kleur (zonder ABGR/ARGB issues).
    param($Cell, [string]$HexColor)
    try {
        $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$HexColor"))
    } catch {
        # Geen harde fout: opmaak is "nice to have", data mag nooit verloren gaan hierdoor.
        Write-Log "Kon celkleur niet zetten ($HexColor): $($_.Exception.Message)" -Level Warning
    }
}

function Test-RequiredModule {
    param([string[]]$Name)
    foreach ($m in $Name) {
        if (-not (Get-Module -ListAvailable -Name $m)) {
            Write-Log "Module '$m' niet gevonden, installeren voor huidige gebruiker..." -Level Warning
            try {
                Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            } catch {
                throw "Kan module '$m' niet automatisch installeren: $($_.Exception.Message). Installeer handmatig met: Install-Module $m -Scope CurrentUser"
            }
        }
        Import-Module -Name $m -ErrorAction Stop
    }
}

function Find-LatestVisioExportCsv {
    param([string]$Folder = (Join-Path $env:USERPROFILE 'Downloads'))
    if (-not (Test-Path $Folder)) { return $null }
    Get-ChildItem -Path $Folder -Filter '*.csv' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match 'visio' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}

function Write-TableToWorksheet {
    # Schrijft een array PSCustomObjects als tabel (header + rijen) handmatig weg.
    # Vermijdt de generieke LoadFromCollection-overload die op PSCustomObject[] faalt
    # bij bepaalde EPPlus/ImportExcel-versies. Geeft start-/eindrij terug.
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][int]$StartRow,
        [int]$StartCol = 1,
        [Parameter(Mandatory)]$Data
    )
    $items = @($Data)
    if ($items.Count -eq 0) {
        return [PSCustomObject]@{ HeaderRow = $StartRow; EndRow = $StartRow; StartCol = $StartCol; EndCol = $StartCol }
    }
    $props = @($items[0].PSObject.Properties.Name)

    for ($c = 0; $c -lt $props.Count; $c++) {
        $colIdx = $StartCol + $c
        $Worksheet.Cells[$StartRow, $colIdx].Value = $props[$c]
    }
    for ($r = 0; $r -lt $items.Count; $r++) {
        $rowIdx = $StartRow + 1 + $r
        for ($c = 0; $c -lt $props.Count; $c++) {
            $colIdx = $StartCol + $c
            $Worksheet.Cells[$rowIdx, $colIdx].Value = $items[$r].($props[$c])
        }
    }
    return [PSCustomObject]@{
        HeaderRow = $StartRow
        EndRow    = $StartRow + $items.Count
        StartCol  = $StartCol
        EndCol    = $StartCol + $props.Count - 1
    }
}

#endregion

#region Graph

function Connect-ToGraph {
    param([string]$TenantId)

    $context = Get-MgContext
    if ($context -and (-not $TenantId -or $context.TenantId -eq $TenantId)) {
        Write-Log "Reeds verbonden met Microsoft Graph als $($context.Account) (tenant: $($context.TenantId))."
        return
    }

    $scopes = @('User.Read.All', 'Organization.Read.All')
    Write-Log "Verbinden met Microsoft Graph (scopes: $($scopes -join ', '))..."

    $connectParams = @{ Scopes = $scopes; NoWelcome = $true }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }

    Connect-MgGraph @connectParams
    Write-Log "Verbonden als $((Get-MgContext).Account)."
}

function Get-VisioPlan2LicensedUsers {
    param([string]$SkuPartNumber)

    Write-Log "Visio licentie opzoeken (SkuPartNumber: $SkuPartNumber)..."
    $allSkus = Get-MgSubscribedSku -All

    $visioSku = $allSkus | Where-Object { $_.SkuPartNumber -eq $SkuPartNumber }
    if (-not $visioSku) {
        $candidates = $allSkus | Where-Object { $_.SkuPartNumber -like '*VISIO*' }
        if ($candidates) {
            $list = ($candidates | ForEach-Object { " - $($_.SkuPartNumber)" }) -join "`n"
            throw "SKU '$SkuPartNumber' niet gevonden in deze tenant. Mogelijke alternatieven:`n$list`nGebruik -VisioSkuPartNumber om de juiste op te geven."
        }
        throw "Geen Visio-gerelateerde licentie gevonden in deze tenant."
    }

    Write-Log "Visio SKU gevonden: $($visioSku.SkuPartNumber) ($($visioSku.SkuId)) - $($visioSku.ConsumedUnits) toegewezen eenheden."

    Write-Log "Alle gebruikers ophalen en filteren op toegewezen Visio licentie..."
    $allUsers = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Department, AccountEnabled, AssignedLicenses

    $targetSkuId = $visioSku.SkuId.ToString().ToLower()
    $licensedUsers = $allUsers | Where-Object {
        $skuIds = $_.AssignedLicenses | ForEach-Object { $_.SkuId.ToString().ToLower() }
        $skuIds -contains $targetSkuId
    }

    Write-Log "Aantal gebruikers met Visio Plan 2 licentie: $($licensedUsers.Count)"
    return $licensedUsers
}

#endregion

#region CSV import

function Import-VisioActivityCsv {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path $Path)) { throw "CSV-bestand niet gevonden: $Path" }

    $firstLine = Get-Content -Path $Path -TotalCount 1
    $delimiter = ','
    if (($firstLine -split ';').Count -gt ($firstLine -split ',').Count) { $delimiter = ';' }

    $raw = Import-Csv -Path $Path -Delimiter $delimiter -Encoding UTF8
    if (-not $raw -or $raw.Count -eq 0) { throw "CSV bevat geen gegevens (alleen header of leeg bestand)." }

    $columns = $raw[0].PSObject.Properties.Name

    function Resolve-Col {
        param([string[]]$Patterns)
        foreach ($p in $Patterns) {
            $match = $columns | Where-Object { ($_ -replace '[^a-zA-Z]', '').ToLower() -like "*$p*" } | Select-Object -First 1
            if ($match) { return $match }
        }
        return $null
    }

    $colUPN          = Resolve-Col -Patterns @('userprincipalname', 'username')
    $colDisplay      = Resolve-Col -Patterns @('displayname')
    $colLastActivity = Resolve-Col -Patterns @('lastactivitydate')
    $colDesktop      = Resolve-Col -Patterns @('desktop')
    $colWeb          = Resolve-Col -Patterns @('web')

    if (-not $colUPN -or -not $colLastActivity) {
        throw "Kan vereiste kolommen niet herkennen in de CSV (User Principal Name / Last Activity Date). Gevonden kolommen: $($columns -join ', ')"
    }

    foreach ($row in $raw) {
        $upnValue = $row.$colUPN
        if (-not $upnValue) { continue }

        $lastActivityRaw = $row.$colLastActivity
        $lastActivityDate = $null
        if ($lastActivityRaw -and $lastActivityRaw -notmatch '^(NULL|N/A)?$') {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse($lastActivityRaw, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                $lastActivityDate = $parsed
            }
        }

        [PSCustomObject]@{
            UserPrincipalName = $upnValue.Trim()
            DisplayName       = if ($colDisplay) { $row.$colDisplay } else { $null }
            LastActivityDate  = $lastActivityDate
            DesktopUsed       = if ($colDesktop) { $row.$colDesktop -match '^(y|yes|true|1)$' } else { $null }
            WebUsed           = if ($colWeb) { $row.$colWeb -match '^(y|yes|true|1)$' } else { $null }
        }
    }
}

#endregion

#region Data merge

function Merge-VisioUsageData {
    param(
        [Parameter(Mandatory)]$LicensedUsers,
        [Parameter(Mandatory)]$ActivityData,
        [Parameter(Mandatory)][int]$InactiveThresholdDays
    )

    $activityLookup = @{}
    foreach ($a in $ActivityData) {
        if ($a.UserPrincipalName) { $activityLookup[$a.UserPrincipalName.ToLower()] = $a }
    }

    $today = Get-Date

    $details = foreach ($u in $LicensedUsers) {
        $key = $u.UserPrincipalName.ToLower()
        $activity = $activityLookup[$key]

        $lastActivity = if ($activity) { $activity.LastActivityDate } else { $null }
        $daysSince = if ($lastActivity) { [math]::Round(($today - $lastActivity).TotalDays) } else { $null }

        $status = if (-not $lastActivity) {
            'Nooit gebruikt'
        } elseif ($daysSince -le $InactiveThresholdDays) {
            'Actief'
        } else {
            'Inactief'
        }

        [PSCustomObject]@{
            DisplayName       = if ($u.DisplayName) { $u.DisplayName } else { $u.UserPrincipalName }
            UserPrincipalName = $u.UserPrincipalName
            Department        = if ($u.Department) { $u.Department } else { 'Onbekend' }
            Status            = $status
            LastActivityDate  = $lastActivity
            DaysSinceLastUse  = $daysSince
            DesktopGebruikt   = if ($activity -and $activity.DesktopUsed) { 'Ja' } elseif ($activity) { 'Nee' } else { '-' }
            WebGebruikt       = if ($activity -and $activity.WebUsed) { 'Ja' } elseif ($activity) { 'Nee' } else { '-' }
            AccountEnabled    = $u.AccountEnabled
        }
    }

    $licensedKeys = $LicensedUsers.UserPrincipalName | ForEach-Object { $_.ToLower() }
    $orphanActivity = $ActivityData | Where-Object {
        $_.UserPrincipalName -and ($licensedKeys -notcontains $_.UserPrincipalName.ToLower())
    }

    return @{
        Details        = @($details | Sort-Object @{Expression = { if ($null -eq $_.DaysSinceLastUse) { [int]::MaxValue } else { $_.DaysSinceLastUse } }; Descending = $true }, DisplayName)
        OrphanActivity = @($orphanActivity)
    }
}

#endregion

#region Excel report

function New-VisioUsageExcelReport {
    param(
        [Parameter(Mandatory)]$Details,
        $OrphanActivity,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$InactiveThresholdDays,
        [string]$ReportPeriodLabel,
        [switch]$ShowExcelAfterCreation
    )

    if ($Details.Count -eq 0) { throw "Geen rijen om te rapporteren - Excel-bestand wordt niet aangemaakt." }

    $outFolder = Split-Path -Path $Path -Parent
    if ($outFolder -and -not (Test-Path $outFolder)) { New-Item -Path $outFolder -ItemType Directory -Force | Out-Null }
    if (Test-Path $Path) { Remove-Item -Path $Path -Force }

    $detailsExport = $Details | Select-Object DisplayName, UserPrincipalName, Department, Status, LastActivityDate, DaysSinceLastUse, DesktopGebruikt, WebGebruikt, AccountEnabled

    # ---------- Sheet: Details ----------
    $detailsExport | Export-Excel -Path $Path -WorksheetName 'Details' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'tblVisioDetails' -TableStyle Medium6

    $pkg = Open-ExcelPackage -Path $Path
    $wsDetails = $pkg.Workbook.Worksheets['Details']

    # Datumformaat kolom E (LastActivityDate)
    try { $wsDetails.Column(5).Style.Numberformat.Format = 'dd-mm-yyyy' } catch { }

    # Handmatige kleurcodering (robuuster dan losse conditional-formatting cmdlets)
    $statusColors = @{ 'Actief' = 'C6EFCE'; 'Inactief' = 'FFEB9C'; 'Nooit gebruikt' = 'FFC7CE' }
    for ($i = 0; $i -lt $detailsExport.Count; $i++) {
        $excelRow = $i + 2
        $row = $detailsExport[$i]

        $statusHex = $statusColors[$row.Status]
        if ($statusHex) { Set-CellFill -Cell $wsDetails.Cells[$excelRow, 4] -HexColor $statusHex }

        $days = $row.DaysSinceLastUse
        $daysHex = if ($null -eq $days) { 'F2F2F2' }
                   elseif ($days -le $InactiveThresholdDays) { 'C6EFCE' }
                   elseif ($days -le ($InactiveThresholdDays * 3)) { 'FFEB9C' }
                   else { 'FFC7CE' }
        Set-CellFill -Cell $wsDetails.Cells[$excelRow, 6] -HexColor $daysHex
    }
    try { $wsDetails.Cells.AutoFitColumns() } catch { }

    # ---------- KPI berekeningen ----------
    $statusOrder = 'Actief', 'Inactief', 'Nooit gebruikt'
    $statusCounts = $statusOrder | ForEach-Object {
        $s = $_
        [PSCustomObject]@{ Status = $s; Aantal = (@($Details | Where-Object { $_.Status -eq $s })).Count }
    }

    $totalLicenses = $Details.Count
    $activeCount   = ($statusCounts | Where-Object Status -eq 'Actief').Aantal
    $inactiveCount = ($statusCounts | Where-Object Status -eq 'Inactief').Aantal
    $neverCount    = ($statusCounts | Where-Object Status -eq 'Nooit gebruikt').Aantal
    $utilization   = if ($totalLicenses -gt 0) { [math]::Round(($activeCount / $totalLicenses) * 100, 1) } else { 0 }
    $disabledWithLicense = (@($Details | Where-Object { $_.AccountEnabled -eq $false })).Count

    $deptSummary = $Details | Group-Object Department | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{
            Afdeling   = $_.Name
            Gebruikers = $_.Count
            Actief     = (@($_.Group | Where-Object { $_.Status -eq 'Actief' })).Count
        }
    }

    $top10 = $Details | Where-Object { $_.Status -ne 'Actief' } |
        Sort-Object @{Expression = { if ($null -eq $_.DaysSinceLastUse) { [int]::MaxValue } else { $_.DaysSinceLastUse } }; Descending = $true } |
        Select-Object -First 10 DisplayName, Department, Status, DaysSinceLastUse

    # ---------- Sheet: Dashboard ----------
    $wsDash = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'Dashboard'
    $wsDash.View.ShowGridLines = $false

    # Titelbanner
    $wsDash.Cells['A1:H2'].Merge = $true
    $wsDash.Cells['A1'].Value = 'Visio Plan 2 - Gebruiksoverzicht'
    $wsDash.Cells['A1'].Style.Font.Size = 20
    $wsDash.Cells['A1'].Style.Font.Bold = $true
    $wsDash.Cells['A1'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $wsDash.Cells['A1'].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $wsDash.Cells['A1'].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    Set-CellFill -Cell $wsDash.Cells['A1:H2'] -HexColor '1F4E78'

    $wsDash.Cells['A3:H3'].Merge = $true
    $wsDash.Cells['A3'].Value = "Periode: $ReportPeriodLabel  |  Gegenereerd op: $(Get-Date -Format 'dd-MM-yyyy HH:mm')  |  Inactiviteitsdrempel: $InactiveThresholdDays dagen"
    $wsDash.Cells['A3'].Style.Font.Italic = $true
    $wsDash.Cells['A3'].Style.Font.Size = 10
    $wsDash.Cells['A3'].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

    # KPI tegels
    $kpis = @(
        @{ Label = 'Totaal licenties';       Value = $totalLicenses;     Color = 'BDD7EE' }
        @{ Label = 'Actief';                 Value = $activeCount;       Color = 'C6EFCE' }
        @{ Label = 'Inactief';                Value = $inactiveCount;     Color = 'FFEB9C' }
        @{ Label = 'Nooit gebruikt';          Value = $neverCount;        Color = 'FFC7CE' }
        @{ Label = 'Bezettingsgraad';         Value = "$utilization%";    Color = 'D9D2E9' }
    )
    $col = 2
    foreach ($kpi in $kpis) {
        $letter = Convert-ColumnNumberToLetter -Number $col
        $wsDash.Cells["$letter`5"].Value = $kpi.Label
        $wsDash.Cells["$letter`5"].Style.Font.Size = 10
        $wsDash.Cells["$letter`5"].Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(89, 89, 89))
        $wsDash.Cells["$letter`5"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

        $wsDash.Cells["$letter`6"].Value = $kpi.Value
        $wsDash.Cells["$letter`6"].Style.Font.Size = 20
        $wsDash.Cells["$letter`6"].Style.Font.Bold = $true
        $wsDash.Cells["$letter`6"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $wsDash.Cells["$letter`6"].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)
        Set-CellFill -Cell $wsDash.Cells["$letter`6"] -HexColor $kpi.Color

        $wsDash.Column($col).Width = 16
        $col++
    }
    if ($disabledWithLicense -gt 0) {
        $wsDash.Cells['A4'].Value = "Let op: $disabledWithLicense gebruiker(s) met deze licentie hebben een uitgeschakeld account."
        $wsDash.Cells['A4'].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
        $wsDash.Cells['A4'].Style.Font.Italic = $true
    }

    # Statusverdeling tabel + grafiek
    $rowStatus = 9
    $wsDash.Cells["A$rowStatus"].Value = 'Statusverdeling'
    $wsDash.Cells["A$rowStatus"].Style.Font.Bold = $true
    $wsDash.Cells["A$rowStatus"].Style.Font.Size = 13

    $statusRange = Write-TableToWorksheet -Worksheet $wsDash -StartRow ($rowStatus + 1) -StartCol 1 -Data $statusCounts
    $statusHeaderRow = $statusRange.HeaderRow
    $statusEndRow = $statusRange.EndRow
    $wsDash.Cells[$statusHeaderRow, 1, $statusHeaderRow, 2].Style.Font.Bold = $true
    for ($r = $statusHeaderRow + 1; $r -le $statusEndRow; $r++) {
        $st = $wsDash.Cells[$r, 1].Value
        $hex = $statusColors[$st]
        if ($hex) { Set-CellFill -Cell $wsDash.Cells[$r, 1, $r, 2] -HexColor $hex }
    }
    $wsDash.Column(1).Width = 18
    $wsDash.Column(2).Width = 10

    try {
        Add-ExcelChart -Worksheet $wsDash -ChartType PieExploded3D -Title 'Statusverdeling Visio Plan 2' `
            -XRange "A$($statusHeaderRow + 1):A$statusEndRow" -YRange "B$($statusHeaderRow + 1):B$statusEndRow" `
            -Row ($statusEndRow + 1) -Column 0 -Width 420 -Height 280
    } catch {
        Write-Log "Kon statuscirkeldiagram niet aanmaken (mogelijk afwijkende ImportExcel-versie): $($_.Exception.Message)" -Level Warning
    }

    # Top 10 langst inactieve gebruikers
    $rowTop10 = 40
    $wsDash.Cells["A$rowTop10"].Value = 'Top 10 langst inactieve gebruikers'
    $wsDash.Cells["A$rowTop10"].Style.Font.Bold = $true
    $wsDash.Cells["A$rowTop10"].Style.Font.Size = 13

    if ($top10.Count -gt 0) {
        $top10Range = Write-TableToWorksheet -Worksheet $wsDash -StartRow ($rowTop10 + 1) -StartCol 1 -Data $top10
        $t10HeaderRow = $top10Range.HeaderRow
        $t10EndRow = $top10Range.EndRow
        $wsDash.Cells[$t10HeaderRow, 1, $t10HeaderRow, 4].Style.Font.Bold = $true
        for ($r = $t10HeaderRow + 1; $r -le $t10EndRow; $r++) {
            $d = $wsDash.Cells[$r, 4].Value
            $hex = if ($null -eq $d) { 'F2F2F2' }
                   elseif ([double]$d -le $InactiveThresholdDays) { 'C6EFCE' }
                   elseif ([double]$d -le ($InactiveThresholdDays * 3)) { 'FFEB9C' }
                   else { 'FFC7CE' }
            Set-CellFill -Cell $wsDash.Cells[$r, 4] -HexColor $hex
        }
        for ($c = 1; $c -le 4; $c++) { $wsDash.Column($c).Width = 20 }
    } else {
        $wsDash.Cells["A$($rowTop10 + 1)"].Value = '(Geen inactieve gebruikers binnen deze selectie)'
    }

    # Gebruik per afdeling
    $rowDept = 58
    $wsDash.Cells["A$rowDept"].Value = 'Gebruik per afdeling'
    $wsDash.Cells["A$rowDept"].Style.Font.Bold = $true
    $wsDash.Cells["A$rowDept"].Style.Font.Size = 13

    if ($deptSummary.Count -gt 0) {
        $deptRange = Write-TableToWorksheet -Worksheet $wsDash -StartRow ($rowDept + 1) -StartCol 1 -Data $deptSummary
        $deptHeaderRow = $deptRange.HeaderRow
        $deptEndRow = $deptRange.EndRow
        $wsDash.Cells[$deptHeaderRow, 1, $deptHeaderRow, 3].Style.Font.Bold = $true

        try {
            Add-ExcelChart -Worksheet $wsDash -ChartType ColumnClustered -Title 'Visio gebruik per afdeling' `
                -XRange "A$($deptHeaderRow + 1):A$deptEndRow" -YRange "B$($deptHeaderRow + 1):C$deptEndRow" `
                -Row ($deptEndRow + 1) -Column 0 -Width 500 -Height 300
        } catch {
            Write-Log "Kon afdelingsgrafiek niet aanmaken (mogelijk afwijkende ImportExcel-versie): $($_.Exception.Message)" -Level Warning
        }
    }

    # ---------- Sheet: Overig (activiteit zonder Plan 2 licentie) ----------
    if ($OrphanActivity -and $OrphanActivity.Count -gt 0) {
        $orphanExport = $OrphanActivity | Select-Object UserPrincipalName, DisplayName, LastActivityDate, DesktopUsed, WebUsed
        $wsOrphan = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'Overig'
        $wsOrphan.Cells['A1'].Value = 'Visio-activiteit gevonden bij gebruikers zonder (huidige) Visio Plan 2 licentie'
        $wsOrphan.Cells['A1'].Style.Font.Bold = $true
        $wsOrphan.Cells['A1:E1'].Merge = $true
        $orphanRange = Write-TableToWorksheet -Worksheet $wsOrphan -StartRow 3 -StartCol 1 -Data $orphanExport
        $wsOrphan.Cells[$orphanRange.HeaderRow, 1, $orphanRange.HeaderRow, 5].Style.Font.Bold = $true
        try { $wsOrphan.Cells.AutoFitColumns() } catch { }
    }

    # Dashboard als eerste tabblad tonen
    try { $pkg.Workbook.Worksheets.MoveToStart('Dashboard') } catch { }

    Close-ExcelPackage -ExcelPackage $pkg -Show:$ShowExcelAfterCreation
}

#endregion

#region Main

$ErrorActionPreference = 'Stop'

try {
    if (-not $OutputPath) {
        $reportsFolder = Join-Path $PSScriptRoot 'Reports'
        if (-not (Test-Path $reportsFolder)) { New-Item -Path $reportsFolder -ItemType Directory -Force | Out-Null }
        $OutputPath = Join-Path $reportsFolder "Visio_Plan2_Gebruiksoverzicht_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    }
    $script:LogFilePath = [System.IO.Path]::ChangeExtension($OutputPath, '.log')

    Write-Log "=== Start Visio Plan 2 gebruiksrapportage ==="
    Write-Log "Output bestand: $OutputPath"

    if (-not $VisioActivityCsvPath) {
        Write-Log "Geen -VisioActivityCsvPath opgegeven, zoeken naar recente export in Downloads-map..."
        $VisioActivityCsvPath = Find-LatestVisioExportCsv
        if (-not $VisioActivityCsvPath) {
            throw @'
Geen Visio activiteitenrapport (CSV) gevonden.

Microsoft Graph biedt GEEN API voor Visio gebruiksdata - dit rapport moet handmatig
worden geexporteerd vanuit het Microsoft 365 admin center:

  1. Ga naar https://admin.microsoft.com
  2. Navigeer naar Rapporten (Reports) > Gebruik (Usage)
  3. Selecteer "Visio" in de lijst met rapporten
  4. Kies de gewenste periode (7/30/90/180 dagen)
  5. Klik op "Exporteren" om de CSV te downloaden
  6. Geef het pad mee via -VisioActivityCsvPath, of zet het bestand in je
     Downloads-map met "visio" in de bestandsnaam.
'@
        }
        Write-Log "Gevonden: $VisioActivityCsvPath"
    }

    if (-not (Test-Path $VisioActivityCsvPath)) {
        throw "Het opgegeven CSV-bestand bestaat niet: $VisioActivityCsvPath"
    }

    Write-Log "Vereiste modules controleren..."
    Test-RequiredModule -Name @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'ImportExcel'
    )

    Connect-ToGraph -TenantId $TenantId

    $licensedUsers = Get-VisioPlan2LicensedUsers -SkuPartNumber $VisioSkuPartNumber
    if ($licensedUsers.Count -eq 0) {
        Write-Log "Geen gebruikers gevonden met deze licentie - script stopt." -Level Warning
        return
    }

    Write-Log "CSV met Visio activiteit inlezen: $VisioActivityCsvPath"
    $activityData = Import-VisioActivityCsv -Path $VisioActivityCsvPath
    Write-Log "Aantal activiteitsregels ingelezen: $($activityData.Count)"

    Write-Log "Gegevens combineren en categoriseren (drempel: $InactiveThresholdDays dagen)..."
    $merged = Merge-VisioUsageData -LicensedUsers $licensedUsers -ActivityData $activityData -InactiveThresholdDays $InactiveThresholdDays

    Write-Log "Excel-rapport opbouwen: $OutputPath"
    New-VisioUsageExcelReport -Details $merged.Details -OrphanActivity $merged.OrphanActivity -Path $OutputPath `
        -InactiveThresholdDays $InactiveThresholdDays -ReportPeriodLabel $Period -ShowExcelAfterCreation:$ShowExcelAfterCreation

    Write-Log "=== Klaar ==="
    $merged.Details | Group-Object Status | ForEach-Object { Write-Log (" - {0}: {1}" -f $_.Name, $_.Count) }
    if ($merged.OrphanActivity.Count -gt 0) {
        Write-Log "Let op: $($merged.OrphanActivity.Count) gebruiker(s) met Visio-activiteit maar zonder huidige Plan 2 licentie - zie tabblad 'Overig'." -Level Warning
    }
    Write-Log "Rapport opgeslagen op: $OutputPath"

    if ($OpenOutputFolder) { Invoke-Item -Path (Split-Path -Path $OutputPath -Parent) }
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)" -Level Error
    throw
}

#endregion
