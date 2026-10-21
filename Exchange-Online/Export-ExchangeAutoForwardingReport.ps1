#requires -Version 5.1
<#
.SYNOPSIS
    Genereert een opgemaakt Excel-rapport van actieve Exchange Online auto-forwarding. Versie 1.4.0.

.DESCRIPTION
    Controleert twee vormen van forwarding:
    1. Mailbox-level forwarding via ForwardingAddress / ForwardingSmtpAddress.
    2. Inbox Rules met ForwardTo, RedirectTo of ForwardAsAttachmentTo.

    Het rapport bevat:
    - Dashboard met KPI-tegels en grafieken.
    - Actieve forwardingconfiguraties.
    - Uitgeschakelde forwardingregels.
    - Scanlog met eventuele fouten.
    - Toelichting en scope.

.PARAMETER OutputPath
    Volledig pad van het te genereren XLSX-bestand.

.PARAMETER RecipientTypeDetails
    Mailboxtypen die worden gescand. Standaard UserMailbox en SharedMailbox.

.PARAMETER OpenReport
    Opent het rapport na voltooiing.

.EXAMPLE
    .\Export-ExchangeAutoForwardingReport.ps1 -OpenReport

.EXAMPLE
    .\Export-ExchangeAutoForwardingReport.ps1 `
        -OutputPath 'C:\Temp\Exchange_Forwarding_Audit.xlsx' `
        -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = (Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath ("Exchange_AutoForwarding_Audit_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [Parameter()]
    [ValidateSet(
        'UserMailbox',
        'SharedMailbox',
        'RoomMailbox',
        'EquipmentMailbox',
        'DiscoveryMailbox'
    )]
    [string[]]$RecipientTypeDetails = @('UserMailbox', 'SharedMailbox'),

    [Parameter()]
    [switch]$OpenReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.4.0'

#region Helper functions
function Import-RequiredModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [string]$InstallCommand
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        $message = "De PowerShell-module '$Name' is niet geïnstalleerd."
        if ($InstallCommand) {
            $message += " Installeer deze eerst met: $InstallCommand"
        }
        throw $message
    }

    Import-Module -Name $Name -ErrorAction Stop
}

function ConvertTo-AddressList {
    [CmdletBinding()]
    param(
        [Parameter()]
        $InputObject
    )

    $addresses = foreach ($item in @($InputObject)) {
        if ($null -eq $item) {
            continue
        }

        $candidate = $null
        foreach ($propertyName in @('PrimarySmtpAddress', 'SmtpAddress', 'Address')) {
            $property = $item.PSObject.Properties[$propertyName]
            if ($property -and $property.Value) {
                $candidate = [string]$property.Value
                break
            }
        }

        if ([string]::IsNullOrWhiteSpace($candidate)) {
            $candidate = [string]$item
        }

        $candidate = $candidate.Trim()
        $candidate = $candidate -replace '(?i)^smtp:', ''

        if ($candidate -match '<(?<Email>[^<>\s]+@[^<>\s]+)>') {
            $candidate = $Matches.Email
        }

        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $candidate
        }
    }

    @($addresses | Sort-Object -Unique)
}

function Get-EmailDomain {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Address
    )

    if ([string]::IsNullOrWhiteSpace($Address)) {
        return $null
    }

    $match = [regex]::Match(
        $Address,
        '(?i)(?<Email>[a-z0-9.!#$%&''*+/=?^_`{|}~-]+@[a-z0-9.-]+\.[a-z]{2,})'
    )

    if (-not $match.Success) {
        return $null
    }

    ($match.Groups['Email'].Value.Split('@')[-1]).ToLowerInvariant()
}

function Test-AcceptedDomain {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Domain,

        [Parameter(Mandatory)]
        [string[]]$AcceptedDomains
    )

    if ([string]::IsNullOrWhiteSpace($Domain)) {
        return $false
    }

    foreach ($acceptedDomain in $AcceptedDomains) {
        $acceptedDomain = $acceptedDomain.ToLowerInvariant()

        if ($acceptedDomain.StartsWith('*.')) {
            $baseDomain = $acceptedDomain.Substring(2)
            if ($Domain -eq $baseDomain -or $Domain.EndsWith(".$baseDomain")) {
                return $true
            }
        }
        elseif ($Domain -eq $acceptedDomain) {
            return $true
        }
    }

    return $false
}

function Get-TargetScope {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$Addresses,

        [Parameter(Mandatory)]
        [string[]]$AcceptedDomains
    )

    $scopes = foreach ($address in @($Addresses)) {
        $domain = Get-EmailDomain -Address $address

        if ([string]::IsNullOrWhiteSpace($domain)) {
            'Onbekend'
        }
        elseif (Test-AcceptedDomain -Domain $domain -AcceptedDomains $AcceptedDomains) {
            'Intern'
        }
        else {
            'Extern'
        }
    }

    if ($scopes -contains 'Extern') {
        return 'Extern'
    }

    if ($scopes -contains 'Onbekend') {
        return 'Onbekend'
    }

    return 'Intern'
}

function Resolve-ForwardingAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ForwardingAddress
    )

    try {
        $recipient = Get-Recipient -Identity $ForwardingAddress -ErrorAction Stop
        if ($recipient.PrimarySmtpAddress) {
            return [string]$recipient.PrimarySmtpAddress
        }
    }
    catch {
        # De oorspronkelijke waarde wordt hieronder gebruikt.
    }

    return [string]$ForwardingAddress
}

function New-ForwardingRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Mailbox,

        [Parameter(Mandatory)]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$RuleName,

        [Parameter()]
        [object]$RulePriority,

        [Parameter(Mandatory)]
        [string[]]$Targets,

        [Parameter(Mandatory)]
        [ValidateSet('Intern', 'Extern', 'Onbekend')]
        [string]$DestinationScope,

        [Parameter(Mandatory)]
        [bool]$Enabled,

        [Parameter()]
        [AllowNull()]
        [object]$KeepCopy,

        [Parameter()]
        [AllowNull()]
        [object]$StopProcessingRules,

        [Parameter()]
        [string[]]$From,

        [Parameter()]
        [string]$Description
    )

    $risk = if (-not $Enabled) {
        'Inactief'
    }
    elseif ($DestinationScope -eq 'Extern') {
        'Verhoogd'
    }
    elseif ($DestinationScope -eq 'Onbekend') {
        'Controleren'
    }
    else {
        'Normaal'
    }

    [pscustomobject][ordered]@{
        'Mailbox'              = [string]$Mailbox.DisplayName
        'Primair SMTP-adres'   = [string]$Mailbox.PrimarySmtpAddress
        'UPN'                  = [string]$Mailbox.UserPrincipalName
        'Mailbox type'         = [string]$Mailbox.RecipientTypeDetails
        'Methode'              = $Method
        'Regelnaam'            = $RuleName
        'Prioriteit'           = $RulePriority
        'Doeladres'            = ($Targets -join '; ')
        'Bestemming'           = $DestinationScope
        'Kopie behouden'       = if ($null -eq $KeepCopy) { 'N.v.t.' } elseif ([bool]$KeepCopy) { 'Ja' } else { 'Nee' }
        'Actief'               = if ($Enabled) { 'Ja' } else { 'Nee' }
        'Stop verdere regels'  = if ($null -eq $StopProcessingRules) { 'N.v.t.' } elseif ([bool]$StopProcessingRules) { 'Ja' } else { 'Nee' }
        'Van / voorwaarde'     = (@($From) -join '; ')
        'Beschrijving'         = $Description
        'Risico-indicatie'     = $risk
        'Entra object-id'      = [string]$Mailbox.ExternalDirectoryObjectId
    }
}

function Set-CellFill {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Range,

        [Parameter(Mandatory)]
        [string]$HexColor
    )

    $Range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $Range.Style.Fill.BackgroundColor.SetColor(
        [System.Drawing.ColorTranslator]::FromHtml($HexColor)
    )
}

function Initialize-DashboardCharting {
    [CmdletBinding()]
    param()

    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction Stop
    }
    catch {
        throw ("De Windows-grafiekbibliotheek kon niet worden geladen: {0}" -f $_.Exception.Message)
    }
}

function New-DoughnutChartImage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Data,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [int]$Width = 1120,

        [Parameter()]
        [int]$Height = 640
    )

    Initialize-DashboardCharting

    $plotData = @($Data | Where-Object { [int]$_.Aantal -gt 0 })
    if ($plotData.Count -eq 0) {
        throw 'Er zijn geen waarden groter dan nul om in de doughnutgrafiek te tonen.'
    }

    $chart = [System.Windows.Forms.DataVisualization.Charting.Chart]::new()
    try {
        $chart.Width = $Width
        $chart.Height = $Height
        $chart.BackColor = [System.Drawing.Color]::White
        $chart.AntiAliasing = [System.Windows.Forms.DataVisualization.Charting.AntiAliasingStyles]::All
        $chart.TextAntiAliasingQuality = [System.Windows.Forms.DataVisualization.Charting.TextAntiAliasingQuality]::High
        $chart.BorderlineColor = [System.Drawing.ColorTranslator]::FromHtml('#D0D7DE')
        $chart.BorderlineDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Solid
        $chart.BorderlineWidth = 1

        $chartArea = [System.Windows.Forms.DataVisualization.Charting.ChartArea]::new()
        $chartArea.Name = 'DestinationArea'
        $chartArea.BackColor = [System.Drawing.Color]::White
        $chartArea.Position = [System.Windows.Forms.DataVisualization.Charting.ElementPosition]::new(2, 14, 96, 72)
        $null = $chart.ChartAreas.Add($chartArea)

        $title = [System.Windows.Forms.DataVisualization.Charting.Title]::new()
        $title.Text = 'Forwarding naar intern of extern'
        $title.Font = [System.Drawing.Font]::new('Segoe UI', 17, [System.Drawing.FontStyle]::Bold)
        $title.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#17365D')
        $title.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Top
        $title.Alignment = [System.Drawing.ContentAlignment]::MiddleCenter
        $null = $chart.Titles.Add($title)

        $legend = [System.Windows.Forms.DataVisualization.Charting.Legend]::new()
        $legend.Name = 'DestinationLegend'
        $legend.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Bottom
        $legend.Alignment = [System.Drawing.StringAlignment]::Center
        $legend.BackColor = [System.Drawing.Color]::Transparent
        $legend.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Regular)
        $legend.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#27364A')
        $legend.IsTextAutoFit = $false
        $null = $chart.Legends.Add($legend)

        $series = [System.Windows.Forms.DataVisualization.Charting.Series]::new()
        $series.Name = 'Bestemming'
        $series.ChartArea = $chartArea.Name
        $series.Legend = $legend.Name
        $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut
        $series.IsValueShownAsLabel = $true
        $series.Label = '#VALX: #VAL (#PERCENT{P0})'
        $series.LegendText = '#VALX'
        $series.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $series.LabelForeColor = [System.Drawing.ColorTranslator]::FromHtml('#27364A')
        $series.BorderColor = [System.Drawing.Color]::White
        $series.BorderWidth = 3
        $series['DoughnutRadius'] = '62'
        $series['PieLabelStyle'] = 'Outside'

        foreach ($item in $plotData) {
            $pointIndex = $series.Points.AddXY([string]$item.Bestemming, [double]$item.Aantal)
            $point = $series.Points[$pointIndex]

            switch ([string]$item.Bestemming) {
                'Intern'   { $point.Color = [System.Drawing.ColorTranslator]::FromHtml('#70AD47') }
                'Extern'   { $point.Color = [System.Drawing.ColorTranslator]::FromHtml('#C00000') }
                default    { $point.Color = [System.Drawing.ColorTranslator]::FromHtml('#FFC000') }
            }
        }

        $null = $chart.Series.Add($series)
        $chart.SaveImage($Path, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)
    }
    finally {
        $chart.Dispose()
    }
}

function New-BarChartImage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Data,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [int]$Width = 1300,

        [Parameter()]
        [int]$Height = 640
    )

    Initialize-DashboardCharting

    $plotData = @($Data | Where-Object { [int]$_.Aantal -gt 0 })
    if ($plotData.Count -eq 0) {
        throw 'Er zijn geen waarden groter dan nul om in de staafgrafiek te tonen.'
    }

    $chart = [System.Windows.Forms.DataVisualization.Charting.Chart]::new()
    try {
        $chart.Width = $Width
        $chart.Height = $Height
        $chart.BackColor = [System.Drawing.Color]::White
        $chart.AntiAliasing = [System.Windows.Forms.DataVisualization.Charting.AntiAliasingStyles]::All
        $chart.TextAntiAliasingQuality = [System.Windows.Forms.DataVisualization.Charting.TextAntiAliasingQuality]::High
        $chart.BorderlineColor = [System.Drawing.ColorTranslator]::FromHtml('#D0D7DE')
        $chart.BorderlineDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Solid
        $chart.BorderlineWidth = 1

        $chartArea = [System.Windows.Forms.DataVisualization.Charting.ChartArea]::new()
        $chartArea.Name = 'MethodArea'
        $chartArea.BackColor = [System.Drawing.Color]::White
        $chartArea.Position = [System.Windows.Forms.DataVisualization.Charting.ElementPosition]::new(5, 15, 91, 78)
        $chartArea.AxisX.Interval = 1
        $chartArea.AxisX.IsReversed = $true
        $chartArea.AxisX.MajorGrid.Enabled = $false
        $chartArea.AxisX.LineColor = [System.Drawing.ColorTranslator]::FromHtml('#AEB8C4')
        $chartArea.AxisX.LabelStyle.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Regular)
        $chartArea.AxisX.LabelStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#27364A')
        $chartArea.AxisY.IsStartedFromZero = $true
        $chartArea.AxisY.Minimum = 0
        $chartArea.AxisY.LabelStyle.Format = '0'
        $chartArea.AxisY.LabelStyle.Font = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
        $chartArea.AxisY.LabelStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#52606D')
        $chartArea.AxisY.LineColor = [System.Drawing.ColorTranslator]::FromHtml('#AEB8C4')
        $chartArea.AxisY.MajorGrid.LineColor = [System.Drawing.ColorTranslator]::FromHtml('#E6EBF0')
        $chartArea.AxisY.MajorGrid.LineDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Solid
        $null = $chart.ChartAreas.Add($chartArea)

        $title = [System.Windows.Forms.DataVisualization.Charting.Title]::new()
        $title.Text = 'Actieve forwarding per methode'
        $title.Font = [System.Drawing.Font]::new('Segoe UI', 17, [System.Drawing.FontStyle]::Bold)
        $title.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#17365D')
        $title.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Top
        $title.Alignment = [System.Drawing.ContentAlignment]::MiddleCenter
        $null = $chart.Titles.Add($title)

        $series = [System.Windows.Forms.DataVisualization.Charting.Series]::new()
        $series.Name = 'Methode'
        $series.ChartArea = $chartArea.Name
        $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Bar
        $series.Color = [System.Drawing.ColorTranslator]::FromHtml('#4472C4')
        $series.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#2F5597')
        $series.BorderWidth = 1
        $series.IsValueShownAsLabel = $true
        $series.LabelFormat = '0'
        $series.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $series.LabelForeColor = [System.Drawing.ColorTranslator]::FromHtml('#17365D')
        $series['PointWidth'] = '0.62'

        foreach ($item in $plotData) {
            $pointIndex = $series.Points.AddXY([string]$item.Methode, [double]$item.Aantal)
            $point = $series.Points[$pointIndex]
            $point.ToolTip = ('{0}: {1}' -f $item.Methode, $item.Aantal)
        }

        $null = $chart.Series.Add($series)
        $chart.SaveImage($Path, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)
    }
    finally {
        $chart.Dispose()
    }
}

function Add-WorksheetPicture {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Worksheet,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [int]$Row,

        [Parameter(Mandatory)]
        [int]$Column,

        [Parameter(Mandatory)]
        [int]$Width,

        [Parameter(Mandatory)]
        [int]$Height,

        [Parameter(Mandatory)]
        $ResourceList
    )

    $image = [System.Drawing.Image]::FromFile($Path)
    $ResourceList.Add($image)

    $picture = $Worksheet.Drawings.AddPicture($Name, $image)
    $picture.SetPosition($Row - 1, 4, $Column - 1, 4)
    $picture.SetSize($Width, $Height)
    return $picture
}

function Set-DashboardSummaryPanel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Worksheet,

        [Parameter(Mandatory)]
        [int]$StartRow,

        [Parameter(Mandatory)]
        [int]$EndRow,

        [Parameter(Mandatory)]
        [int]$StartColumn,

        [Parameter(Mandatory)]
        [int]$EndColumn,

        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [object[]]$Data,

        [Parameter(Mandatory)]
        [string]$LabelProperty,

        [Parameter(Mandatory)]
        [string]$ValueProperty
    )

    $titleEndRow = $StartRow + 1
    $innerStartColumn = $StartColumn + 1
    $innerEndColumn = $EndColumn - 1
    $labelEndColumn = $EndColumn - 2
    $valueStartColumn = $labelEndColumn + 1

    $panelRange = $Worksheet.Cells[$StartRow, $StartColumn, $EndRow, $EndColumn]
    $panelRange.Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $panelRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $panelRange.Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $panelRange.Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $panelRange.Style.Border.Top.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#D0D7DE'))
    $panelRange.Style.Border.Bottom.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#D0D7DE'))
    $panelRange.Style.Border.Left.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#D0D7DE'))
    $panelRange.Style.Border.Right.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#D0D7DE'))
    Set-CellFill -Range $panelRange -HexColor '#FFFFFF'

    $titleRange = $Worksheet.Cells[$StartRow, $StartColumn, $titleEndRow, $EndColumn]
    $titleRange.Merge = $true
    $Worksheet.Cells[$StartRow, $StartColumn].Value = $Title
    Set-CellFill -Range $titleRange -HexColor '#17365D'
    $titleRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $titleRange.Style.Font.Bold = $true
    $titleRange.Style.Font.Size = 13
    $titleRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $titleRange.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

    $row = $StartRow + 3
    foreach ($item in $Data) {
        if ($row -gt ($EndRow - 1)) {
            break
        }

        $rowEnd = $row + 1
        $labelRange = $Worksheet.Cells[$row, $innerStartColumn, $rowEnd, $labelEndColumn]
        $valueRange = $Worksheet.Cells[$row, $valueStartColumn, $rowEnd, $innerEndColumn]
        $dataRange = $Worksheet.Cells[$row, $innerStartColumn, $rowEnd, $innerEndColumn]

        $labelRange.Merge = $true
        $valueRange.Merge = $true
        $Worksheet.Cells[$row, $innerStartColumn].Value = [string]$item.PSObject.Properties[$LabelProperty].Value
        $Worksheet.Cells[$row, $valueStartColumn].Value = [int]$item.PSObject.Properties[$ValueProperty].Value

        $stripeIndex = [int](($row - ($StartRow + 3)) / 2)
        if (($stripeIndex % 2) -eq 0) {
            Set-CellFill -Range $dataRange -HexColor '#F4F7FA'
        }
        else {
            Set-CellFill -Range $dataRange -HexColor '#EAF2F8'
        }

        $dataRange.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $labelRange.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#27364A'))
        $labelRange.Style.Font.Bold = $true
        $valueRange.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#17365D'))
        $valueRange.Style.Font.Bold = $true
        $valueRange.Style.Font.Size = 14
        $valueRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

        $row += 2
    }
}

function Set-KpiCard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Worksheet,

        [Parameter(Mandatory)]
        [string]$LabelRange,

        [Parameter(Mandatory)]
        [string]$ValueRange,

        [Parameter(Mandatory)]
        [string]$Label,

        [Parameter(Mandatory)]
        $Value,

        [Parameter(Mandatory)]
        [string]$BackgroundColor
    )

    $Worksheet.Cells[$LabelRange].Merge = $true
    $Worksheet.Cells[$ValueRange].Merge = $true

    $Worksheet.Cells[$LabelRange].Value = $Label.ToUpperInvariant()
    $Worksheet.Cells[$ValueRange].Value = $Value

    foreach ($rangeName in @($LabelRange, $ValueRange)) {
        $range = $Worksheet.Cells[$rangeName]
        Set-CellFill -Range $range -HexColor $BackgroundColor
        $range.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $range.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $range.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $range.Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $range.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $range.Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $range.Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $range.Style.Border.Top.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#D7E3F4'))
        $range.Style.Border.Bottom.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#D7E3F4'))
        $range.Style.Border.Left.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#D7E3F4'))
        $range.Style.Border.Right.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#D7E3F4'))
    }

    $Worksheet.Cells[$LabelRange].Style.Font.Size = 10
    $Worksheet.Cells[$LabelRange].Style.Font.Bold = $true
    $Worksheet.Cells[$ValueRange].Style.Font.Size = 26
    $Worksheet.Cells[$ValueRange].Style.Font.Bold = $true
}

function Format-DataWorksheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Worksheet,

        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [string]$Subtitle,

        [Parameter(Mandatory)]
        [int]$DataRowCount
    )

    $lastColumn = 16
    $lastDataRow = [Math]::Max(5, 4 + $DataRowCount)

    $Worksheet.Cells[1, 1, 1, $lastColumn].Merge = $true
    $Worksheet.Cells[1, 1].Value = $Title
    Set-CellFill -Range $Worksheet.Cells[1, 1, 1, $lastColumn] -HexColor '#17365D'
    $Worksheet.Cells[1, 1, 1, $lastColumn].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $Worksheet.Cells[1, 1, 1, $lastColumn].Style.Font.Bold = $true
    $Worksheet.Cells[1, 1, 1, $lastColumn].Style.Font.Size = 20
    $Worksheet.Cells[1, 1, 1, $lastColumn].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $Worksheet.Row(1).Height = 34

    $Worksheet.Cells[2, 1, 2, $lastColumn].Merge = $true
    $Worksheet.Cells[2, 1].Value = $Subtitle
    Set-CellFill -Range $Worksheet.Cells[2, 1, 2, $lastColumn] -HexColor '#D9EAF7'
    $Worksheet.Cells[2, 1, 2, $lastColumn].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#17365D'))
    $Worksheet.Cells[2, 1, 2, $lastColumn].Style.Font.Italic = $true
    $Worksheet.Cells[2, 1, 2, $lastColumn].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $Worksheet.Row(2).Height = 24

    $Worksheet.View.ShowGridLines = $false
    $Worksheet.View.FreezePanes(5, 1)
    $Worksheet.View.ZoomScale = 90

    $widths = @{
        1 = 28; 2 = 32; 3 = 32; 4 = 20; 5 = 31; 6 = 31; 7 = 11; 8 = 44;
        9 = 14; 10 = 15; 11 = 10; 12 = 18; 13 = 34; 14 = 70; 15 = 18; 16 = 39
    }

    foreach ($column in $widths.Keys) {
        $Worksheet.Column($column).Width = $widths[$column]
    }

    $Worksheet.Cells[5, 8, $lastDataRow, 8].Style.WrapText = $true
    $Worksheet.Cells[5, 13, $lastDataRow, 14].Style.WrapText = $true
    $Worksheet.Cells[5, 1, $lastDataRow, $lastColumn].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top

    for ($row = 5; $row -le $lastDataRow; $row++) {
        $destination = [string]$Worksheet.Cells[$row, 9].Value
        $risk = [string]$Worksheet.Cells[$row, 15].Value
        $active = [string]$Worksheet.Cells[$row, 11].Value

        if ($destination -eq 'Extern') {
            Set-CellFill -Range $Worksheet.Cells[$row, 9] -HexColor '#FCE4D6'
            $Worksheet.Cells[$row, 9].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#9C0006'))
            $Worksheet.Cells[$row, 9].Style.Font.Bold = $true
        }
        elseif ($destination -eq 'Intern') {
            Set-CellFill -Range $Worksheet.Cells[$row, 9] -HexColor '#E2F0D9'
            $Worksheet.Cells[$row, 9].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#375623'))
        }

        switch ($risk) {
            'Verhoogd' {
                Set-CellFill -Range $Worksheet.Cells[$row, 15] -HexColor '#F4CCCC'
                $Worksheet.Cells[$row, 15].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#9C0006'))
                $Worksheet.Cells[$row, 15].Style.Font.Bold = $true
            }
            'Controleren' {
                Set-CellFill -Range $Worksheet.Cells[$row, 15] -HexColor '#FFF2CC'
                $Worksheet.Cells[$row, 15].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#7F6000'))
            }
            'Normaal' {
                Set-CellFill -Range $Worksheet.Cells[$row, 15] -HexColor '#E2F0D9'
                $Worksheet.Cells[$row, 15].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#375623'))
            }
        }

        if ($active -eq 'Nee') {
            $Worksheet.Cells[$row, 1, $row, $lastColumn].Style.Font.Color.SetColor(
                [System.Drawing.ColorTranslator]::FromHtml('#7F7F7F')
            )
        }
    }
}
#endregion Helper functions

#region Prerequisites and connection
Import-RequiredModule -Name 'ExchangeOnlineManagement' -InstallCommand 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'
Import-RequiredModule -Name 'ImportExcel' -InstallCommand 'Install-Module ImportExcel -Scope CurrentUser'

$outputDirectory = Split-Path -Path $OutputPath -Parent
if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
    $outputDirectory = (Get-Location).Path
    $OutputPath = Join-Path -Path $outputDirectory -ChildPath $OutputPath
}

if (-not (Test-Path -LiteralPath $outputDirectory)) {
    $null = New-Item -Path $outputDirectory -ItemType Directory -Force
}

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

Write-Host ("Exchange Auto-Forwarding Report v{0}" -f $ScriptVersion) -ForegroundColor Magenta
Write-Host 'Verbinden met Exchange Online...' -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false
#endregion Prerequisites and connection

$scanStarted = Get-Date
$records = [System.Collections.Generic.List[object]]::new()
$scanLog = [System.Collections.Generic.List[object]]::new()
$dashboardImageResources = [System.Collections.Generic.List[System.IDisposable]]::new()
$dashboardAssetDirectory = $null

try {
    Write-Host 'Accepted domains ophalen...' -ForegroundColor Cyan
    $acceptedDomains = @(
        Get-AcceptedDomain -ErrorAction Stop |
            ForEach-Object { $_.DomainName.ToString().ToLowerInvariant() }
    )

    Write-Host 'Mailboxen ophalen...' -ForegroundColor Cyan

    # Gebruik expliciet een string-array. Dit voorkomt typeproblemen wanneer slechts
    # één mailboxtype wordt meegegeven of wanneer een moduleversie parameters anders bindt.
    [string[]]$recipientTypes = @(
        $RecipientTypeDetails |
            ForEach-Object { [string]$_ }
    )

    # Gebruik Get-Mailbox rechtstreeks. Sommige ExchangeOnlineManagement-versies
    # veroorzaken intern een Object[] + Object[]-fout wanneer Get-EXOMailbox wordt
    # aangeroepen met RecipientTypeDetails en propertysets. Door elk mailboxtype
    # afzonderlijk op te halen wordt geen array aan de cmdlet doorgegeven.
    $mailboxBuffer = [System.Collections.Generic.List[object]]::new()

    foreach ($recipientType in $recipientTypes) {
        Write-Host ("  - {0}" -f $recipientType) -ForegroundColor DarkCyan

        $typeMailboxes = @(
            Get-Mailbox `
                -ResultSize Unlimited `
                -RecipientTypeDetails $recipientType `
                -ErrorAction Stop |
                Select-Object `
                    DisplayName,
                    UserPrincipalName,
                    PrimarySmtpAddress,
                    RecipientTypeDetails,
                    ForwardingAddress,
                    ForwardingSmtpAddress,
                    DeliverToMailboxAndForward,
                    ExternalDirectoryObjectId
        )

        foreach ($typeMailbox in $typeMailboxes) {
            $mailboxBuffer.Add($typeMailbox)
        }
    }

    $mailboxes = @(
        $mailboxBuffer |
            Sort-Object PrimarySmtpAddress -Unique
    )

    if ($mailboxes.Count -eq 0) {
        throw ("Geen mailboxen gevonden voor RecipientTypeDetails: {0}" -f ($recipientTypes -join ', '))
    }

    $mailboxIndex = 0
    foreach ($mailbox in $mailboxes) {
        $mailboxIndex++
        $primarySmtpAddress = [string]$mailbox.PrimarySmtpAddress

        Write-Progress `
            -Activity 'Exchange Online forwarding controleren' `
            -Status ("{0} van {1}: {2}" -f $mailboxIndex, $mailboxes.Count, $primarySmtpAddress) `
            -PercentComplete (($mailboxIndex / [Math]::Max(1, $mailboxes.Count)) * 100)

        # Mailbox-level forwarding: externe SMTP-target.
        if ($mailbox.ForwardingSmtpAddress) {
            $targets = ConvertTo-AddressList -InputObject $mailbox.ForwardingSmtpAddress
            $scope = Get-TargetScope -Addresses $targets -AcceptedDomains $acceptedDomains

            $records.Add((New-ForwardingRecord `
                -Mailbox $mailbox `
                -Method 'Mailbox forwarding - SMTP' `
                -RuleName 'Mailboxinstelling' `
                -Targets $targets `
                -DestinationScope $scope `
                -Enabled $true `
                -KeepCopy $mailbox.DeliverToMailboxAndForward `
                -Description 'Forwarding ingesteld op mailboxniveau via ForwardingSmtpAddress.'))
        }

        # Mailbox-level forwarding: interne recipient-target.
        if ($mailbox.ForwardingAddress) {
            $resolvedAddress = Resolve-ForwardingAddress -ForwardingAddress $mailbox.ForwardingAddress
            $targets = ConvertTo-AddressList -InputObject $resolvedAddress
            $scope = Get-TargetScope -Addresses $targets -AcceptedDomains $acceptedDomains

            $records.Add((New-ForwardingRecord `
                -Mailbox $mailbox `
                -Method 'Mailbox forwarding - recipient' `
                -RuleName 'Mailboxinstelling' `
                -Targets $targets `
                -DestinationScope $scope `
                -Enabled $true `
                -KeepCopy $mailbox.DeliverToMailboxAndForward `
                -Description 'Forwarding ingesteld op mailboxniveau via ForwardingAddress.'))
        }

        # Inbox Rules.
        try {
            $rules = @(
                # ResultSize bewust niet meegeven. De standaardlimiet is 1000 regels
                # per mailbox en voorkomt een parameterbindingfout in sommige EXO-moduleversies.
                Get-InboxRule `
                    -Mailbox $primarySmtpAddress `
                    -IncludeHidden `
                    -ErrorAction Stop
            )

            foreach ($rule in $rules) {
                $actions = @(
                    [pscustomobject]@{ Method = 'Inboxregel - ForwardTo';            Targets = ConvertTo-AddressList -InputObject $rule.ForwardTo }
                    [pscustomobject]@{ Method = 'Inboxregel - RedirectTo';           Targets = ConvertTo-AddressList -InputObject $rule.RedirectTo }
                    [pscustomobject]@{ Method = 'Inboxregel - ForwardAsAttachment';  Targets = ConvertTo-AddressList -InputObject $rule.ForwardAsAttachmentTo }
                )

                foreach ($action in $actions) {
                    if (@($action.Targets).Count -eq 0) {
                        continue
                    }

                    $scope = Get-TargetScope -Addresses $action.Targets -AcceptedDomains $acceptedDomains
                    $from = ConvertTo-AddressList -InputObject $rule.From

                    $records.Add((New-ForwardingRecord `
                        -Mailbox $mailbox `
                        -Method $action.Method `
                        -RuleName ([string]$rule.Name) `
                        -RulePriority $rule.Priority `
                        -Targets $action.Targets `
                        -DestinationScope $scope `
                        -Enabled ([bool]$rule.Enabled) `
                        -StopProcessingRules $rule.StopProcessingRules `
                        -From $from `
                        -Description ([string]$rule.Description)))
                }
            }

            $scanLog.Add([pscustomobject][ordered]@{
                'Mailbox'            = $mailbox.DisplayName
                'Primair SMTP-adres' = $primarySmtpAddress
                'Resultaat'          = 'Succesvol'
                'Melding'            = ''
            })
        }
        catch {
            $scanLog.Add([pscustomobject][ordered]@{
                'Mailbox'            = $mailbox.DisplayName
                'Primair SMTP-adres' = $primarySmtpAddress
                'Resultaat'          = 'Fout'
                'Melding'            = $_.Exception.Message
            })
        }
    }

    Write-Progress -Activity 'Exchange Online forwarding controleren' -Completed

    $scanCompleted = Get-Date
    $activeRecords = @($records | Where-Object { $_.Actief -eq 'Ja' } | Sort-Object Mailbox, Methode, Regelnaam, Doeladres)
    $disabledRecords = @($records | Where-Object { $_.Actief -eq 'Nee' } | Sort-Object Mailbox, Methode, Regelnaam, Doeladres)
    $errorRecords = @($scanLog | Where-Object { $_.Resultaat -eq 'Fout' })

    $uniqueForwardingMailboxes = @($activeRecords.'Primair SMTP-adres' | Sort-Object -Unique).Count
    $externalRecords = @($activeRecords | Where-Object { $_.Bestemming -eq 'Extern' })
    $internalRecords = @($activeRecords | Where-Object { $_.Bestemming -eq 'Intern' })
    $unknownRecords = @($activeRecords | Where-Object { $_.Bestemming -eq 'Onbekend' })

    $methodSummary = @(
        [pscustomobject]@{ 'Methode' = 'Mailbox forwarding'; 'Aantal' = @($activeRecords | Where-Object { $_.Methode -like 'Mailbox forwarding*' }).Count }
        [pscustomobject]@{ 'Methode' = 'Inboxregel: ForwardTo'; 'Aantal' = @($activeRecords | Where-Object { $_.Methode -eq 'Inboxregel - ForwardTo' }).Count }
        [pscustomobject]@{ 'Methode' = 'Inboxregel: RedirectTo'; 'Aantal' = @($activeRecords | Where-Object { $_.Methode -eq 'Inboxregel - RedirectTo' }).Count }
        [pscustomobject]@{ 'Methode' = 'Inboxregel: als bijlage'; 'Aantal' = @($activeRecords | Where-Object { $_.Methode -eq 'Inboxregel - ForwardAsAttachment' }).Count }
    )

    $destinationSummary = @(
        [pscustomobject]@{ 'Bestemming' = 'Intern'; 'Aantal' = $internalRecords.Count }
        [pscustomobject]@{ 'Bestemming' = 'Extern'; 'Aantal' = $externalRecords.Count }
        [pscustomobject]@{ 'Bestemming' = 'Onbekend'; 'Aantal' = $unknownRecords.Count }
    )

    # Export-Excel heeft minimaal één object nodig om een tabel met kolommen te bouwen.
    $activeExport = if ($activeRecords.Count -gt 0) {
        $activeRecords
    }
    else {
        @([pscustomobject][ordered]@{
            'Mailbox' = 'Geen actieve forwarding gevonden'; 'Primair SMTP-adres' = ''; 'UPN' = ''; 'Mailbox type' = '';
            'Methode' = ''; 'Regelnaam' = ''; 'Prioriteit' = ''; 'Doeladres' = ''; 'Bestemming' = ''; 'Kopie behouden' = '';
            'Actief' = ''; 'Stop verdere regels' = ''; 'Van / voorwaarde' = ''; 'Beschrijving' = '';
            'Risico-indicatie' = ''; 'Entra object-id' = ''
        })
    }

    $disabledExport = if ($disabledRecords.Count -gt 0) {
        $disabledRecords
    }
    else {
        @([pscustomobject][ordered]@{
            'Mailbox' = 'Geen uitgeschakelde forwardingregels gevonden'; 'Primair SMTP-adres' = ''; 'UPN' = ''; 'Mailbox type' = '';
            'Methode' = ''; 'Regelnaam' = ''; 'Prioriteit' = ''; 'Doeladres' = ''; 'Bestemming' = ''; 'Kopie behouden' = '';
            'Actief' = ''; 'Stop verdere regels' = ''; 'Van / voorwaarde' = ''; 'Beschrijving' = '';
            'Risico-indicatie' = ''; 'Entra object-id' = ''
        })
    }

    $excelPackage = $activeExport | Export-Excel `
        -Path $OutputPath `
        -WorksheetName 'Actieve forwarding' `
        -StartRow 4 `
        -TableName 'tblActiveForwarding' `
        -TableStyle Medium2 `
        -AutoFilter `
        -BoldTopRow `
        -PassThru

    $null = $disabledExport | Export-Excel `
        -ExcelPackage $excelPackage `
        -WorksheetName 'Uitgeschakelde regels' `
        -StartRow 4 `
        -TableName 'tblDisabledForwarding' `
        -TableStyle Medium15 `
        -AutoFilter `
        -BoldTopRow `
        -PassThru

    $scanLogExport = if ($scanLog.Count -gt 0) {
        @($scanLog)
    }
    else {
        @([pscustomobject][ordered]@{
            'Mailbox' = 'Geen mailboxen gevonden'
            'Primair SMTP-adres' = ''
            'Resultaat' = 'Geen data'
            'Melding' = 'Er zijn geen mailboxen gevonden binnen de geselecteerde RecipientTypeDetails.'
        })
    }

    $null = $scanLogExport | Export-Excel `
        -ExcelPackage $excelPackage `
        -WorksheetName 'Scanlog' `
        -StartRow 4 `
        -TableName 'tblScanLog' `
        -TableStyle Medium9 `
        -AutoFilter `
        -BoldTopRow `
        -PassThru

    # Data sheets formatteren.
    $wsActive = $excelPackage.Workbook.Worksheets['Actieve forwarding']
    Format-DataWorksheet `
        -Worksheet $wsActive `
        -Title 'ACTIEVE AUTO-FORWARDING' `
        -Subtitle ("Alle actieve mailbox-forwarding en Inbox Rules | Gegenereerd: {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm')) `
        -DataRowCount $activeExport.Count

    $wsDisabled = $excelPackage.Workbook.Worksheets['Uitgeschakelde regels']
    Format-DataWorksheet `
        -Worksheet $wsDisabled `
        -Title 'UITGESCHAKELDE FORWARDINGREGELS' `
        -Subtitle 'Forwardingacties die nog als regel bestaan, maar momenteel niet actief zijn.' `
        -DataRowCount $disabledExport.Count

    # Scanlog formatteren.
    $wsLog = $excelPackage.Workbook.Worksheets['Scanlog']
    $wsLog.Cells['A1:D1'].Merge = $true
    $wsLog.Cells['A1'].Value = 'SCANLOG'
    Set-CellFill -Range $wsLog.Cells['A1:D1'] -HexColor '#17365D'
    $wsLog.Cells['A1:D1'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $wsLog.Cells['A1:D1'].Style.Font.Bold = $true
    $wsLog.Cells['A1:D1'].Style.Font.Size = 20
    $wsLog.Row(1).Height = 34

    $wsLog.Cells['A2:D2'].Merge = $true
    $wsLog.Cells['A2'].Value = ("{0} mailboxen gescand | {1} fout(en) | Doorlooptijd: {2:n1} minuten" -f $mailboxes.Count, $errorRecords.Count, ($scanCompleted - $scanStarted).TotalMinutes)
    Set-CellFill -Range $wsLog.Cells['A2:D2'] -HexColor '#D9EAF7'
    $wsLog.Cells['A2:D2'].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#17365D'))
    $wsLog.View.ShowGridLines = $false
    $wsLog.View.FreezePanes(5, 1)
    $wsLog.Column(1).Width = 30
    $wsLog.Column(2).Width = 38
    $wsLog.Column(3).Width = 15
    $wsLog.Column(4).Width = 90
    $wsLog.Cells[5, 4, [Math]::Max(5, 4 + $scanLogExport.Count), 4].Style.WrapText = $true

    for ($row = 5; $row -le (4 + $scanLogExport.Count); $row++) {
        if (([string]$wsLog.Cells[$row, 3].Value) -eq 'Fout') {
            Set-CellFill -Range $wsLog.Cells[$row, 3] -HexColor '#F4CCCC'
            $wsLog.Cells[$row, 3].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#9C0006'))
            $wsLog.Cells[$row, 3].Style.Font.Bold = $true
        }
        else {
            Set-CellFill -Range $wsLog.Cells[$row, 3] -HexColor '#E2F0D9'
            $wsLog.Cells[$row, 3].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#375623'))
        }
    }

    # Dashboard maken en vooraan plaatsen.
    $wsDashboard = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName 'Dashboard' -MoveToStart
    $wsDashboard.View.ShowGridLines = $false
    $wsDashboard.View.ZoomScale = 90
    $wsDashboard.Cells['B1:P1'].Merge = $true
    $wsDashboard.Cells['B1'].Value = 'EXCHANGE ONLINE — AUTO-FORWARDING AUDIT'
    Set-CellFill -Range $wsDashboard.Cells['B1:P1'] -HexColor '#17365D'
    $wsDashboard.Cells['B1:P1'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $wsDashboard.Cells['B1:P1'].Style.Font.Bold = $true
    $wsDashboard.Cells['B1:P1'].Style.Font.Size = 22
    $wsDashboard.Cells['B1:P1'].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $wsDashboard.Cells['B1:P1'].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $wsDashboard.Row(1).Height = 42

    $wsDashboard.Cells['B2:P2'].Merge = $true
    $wsDashboard.Cells['B2'].Value = ("Tenantbreed inzicht in mailbox-forwarding en Inbox Rules | Rapportdatum: {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
    Set-CellFill -Range $wsDashboard.Cells['B2:P2'] -HexColor '#D9EAF7'
    $wsDashboard.Cells['B2:P2'].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#17365D'))
    $wsDashboard.Cells['B2:P2'].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $wsDashboard.Row(2).Height = 24

    Set-KpiCard -Worksheet $wsDashboard -LabelRange 'B4:D4'  -ValueRange 'B5:D7'  -Label 'Mailboxen gescand'        -Value $mailboxes.Count             -BackgroundColor '#4472C4'
    Set-KpiCard -Worksheet $wsDashboard -LabelRange 'F4:H4'  -ValueRange 'F5:H7'  -Label 'Mailboxen met forwarding' -Value $uniqueForwardingMailboxes  -BackgroundColor '#5B9BD5'
    Set-KpiCard -Worksheet $wsDashboard -LabelRange 'J4:L4'  -ValueRange 'J5:L7'  -Label 'Actieve configuraties'    -Value $activeRecords.Count        -BackgroundColor '#70AD47'
    Set-KpiCard -Worksheet $wsDashboard -LabelRange 'N4:P4'  -ValueRange 'N5:P7'  -Label 'Externe forwarding'       -Value $externalRecords.Count      -BackgroundColor '#C00000'

    # Dashboardgrafieken worden als hoge-resolutie PNG ingebed.
    # Hiermee wordt de versiespecifieke native-chartfout van ImportExcel/EPPlus
    # omzeild waarbij wel een grafiekcontainer, maar geen dataserie wordt opgeslagen.
    $dashboardAssetDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("ExchangeForwardingDashboard_{0}" -f ([guid]::NewGuid().ToString('N')))
    $null = New-Item -Path $dashboardAssetDirectory -ItemType Directory -Force

    $destinationTotal = ($destinationSummary | Measure-Object -Property Aantal -Sum).Sum
    if ($destinationTotal -gt 0) {
        try {
            $destinationChartPath = Join-Path -Path $dashboardAssetDirectory -ChildPath 'forwarding-bestemming.png'
            New-DoughnutChartImage -Data $destinationSummary -Path $destinationChartPath
            $null = Add-WorksheetPicture `
                -Worksheet $wsDashboard `
                -Path $destinationChartPath `
                -Name 'ForwardingDestinationChart' `
                -Row 9 `
                -Column 2 `
                -Width 560 `
                -Height 320 `
                -ResourceList $dashboardImageResources
        }
        catch {
            Write-Warning ("De bestemmingsgrafiek kon niet als afbeelding worden gemaakt. Er wordt een samenvattingspaneel gebruikt. Fout: {0}" -f $_.Exception.Message)
            Set-DashboardSummaryPanel `
                -Worksheet $wsDashboard `
                -StartRow 9 `
                -EndRow 22 `
                -StartColumn 2 `
                -EndColumn 8 `
                -Title 'FORWARDING NAAR INTERN OF EXTERN' `
                -Data $destinationSummary `
                -LabelProperty 'Bestemming' `
                -ValueProperty 'Aantal'
        }
    }
    else {
        Set-DashboardSummaryPanel `
            -Worksheet $wsDashboard `
            -StartRow 9 `
            -EndRow 22 `
            -StartColumn 2 `
            -EndColumn 8 `
            -Title 'FORWARDING NAAR INTERN OF EXTERN' `
            -Data $destinationSummary `
            -LabelProperty 'Bestemming' `
            -ValueProperty 'Aantal'
    }

    $methodTotal = ($methodSummary | Measure-Object -Property Aantal -Sum).Sum
    if ($methodTotal -gt 0) {
        try {
            $methodChartPath = Join-Path -Path $dashboardAssetDirectory -ChildPath 'forwarding-methode.png'
            New-BarChartImage -Data $methodSummary -Path $methodChartPath
            $null = Add-WorksheetPicture `
                -Worksheet $wsDashboard `
                -Path $methodChartPath `
                -Name 'ForwardingMethodChart' `
                -Row 9 `
                -Column 10 `
                -Width 650 `
                -Height 320 `
                -ResourceList $dashboardImageResources
        }
        catch {
            Write-Warning ("De methodegrafiek kon niet als afbeelding worden gemaakt. Er wordt een samenvattingspaneel gebruikt. Fout: {0}" -f $_.Exception.Message)
            Set-DashboardSummaryPanel `
                -Worksheet $wsDashboard `
                -StartRow 9 `
                -EndRow 22 `
                -StartColumn 10 `
                -EndColumn 16 `
                -Title 'ACTIEVE FORWARDING PER METHODE' `
                -Data $methodSummary `
                -LabelProperty 'Methode' `
                -ValueProperty 'Aantal'
        }
    }
    else {
        Set-DashboardSummaryPanel `
            -Worksheet $wsDashboard `
            -StartRow 9 `
            -EndRow 22 `
            -StartColumn 10 `
            -EndColumn 16 `
            -Title 'ACTIEVE FORWARDING PER METHODE' `
            -Data $methodSummary `
            -LabelProperty 'Methode' `
            -ValueProperty 'Aantal'
    }

    # Externe forwarding spotlight.
    $wsDashboard.Cells['B25:P25'].Merge = $true
    $wsDashboard.Cells['B25'].Value = 'EXTERNE FORWARDING — DIRECTE AANDACHT'
    Set-CellFill -Range $wsDashboard.Cells['B25:P25'] -HexColor '#C00000'
    $wsDashboard.Cells['B25:P25'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $wsDashboard.Cells['B25:P25'].Style.Font.Bold = $true
    $wsDashboard.Cells['B25:P25'].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

    $externalSpotlight = @($externalRecords | Select-Object -First 10)
    $spotlightHeaders = @('Mailbox', 'Primair SMTP-adres', 'Methode', 'Doeladres', 'Risico')
    $spotlightColumns = @(2, 5, 9, 12, 16)

    for ($index = 0; $index -lt $spotlightHeaders.Count; $index++) {
        $startColumn = $spotlightColumns[$index]
        $endColumn = if ($index -lt ($spotlightColumns.Count - 1)) { $spotlightColumns[$index + 1] - 1 } else { 16 }
        $wsDashboard.Cells[26, $startColumn, 26, $endColumn].Merge = $true
        $wsDashboard.Cells[26, $startColumn].Value = $spotlightHeaders[$index]
        Set-CellFill -Range $wsDashboard.Cells[26, $startColumn, 26, $endColumn] -HexColor '#D9E1F2'
        $wsDashboard.Cells[26, $startColumn, 26, $endColumn].Style.Font.Bold = $true
        $wsDashboard.Cells[26, $startColumn, 26, $endColumn].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    }

    if ($externalSpotlight.Count -eq 0) {
        $wsDashboard.Cells['B27:P28'].Merge = $true
        $wsDashboard.Cells['B27'].Value = 'Geen actieve externe forwarding gevonden.'
        Set-CellFill -Range $wsDashboard.Cells['B27:P28'] -HexColor '#E2F0D9'
        $wsDashboard.Cells['B27:P28'].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#375623'))
        $wsDashboard.Cells['B27:P28'].Style.Font.Bold = $true
        $wsDashboard.Cells['B27:P28'].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $wsDashboard.Cells['B27:P28'].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    }
    else {
        $row = 27
        foreach ($item in $externalSpotlight) {
            $values = @($item.Mailbox, $item.'Primair SMTP-adres', $item.Methode, $item.Doeladres, $item.'Risico-indicatie')
            for ($index = 0; $index -lt $values.Count; $index++) {
                $startColumn = $spotlightColumns[$index]
                $endColumn = if ($index -lt ($spotlightColumns.Count - 1)) { $spotlightColumns[$index + 1] - 1 } else { 16 }
                $wsDashboard.Cells[$row, $startColumn, $row, $endColumn].Merge = $true
                $wsDashboard.Cells[$row, $startColumn].Value = [string]$values[$index]
                $wsDashboard.Cells[$row, $startColumn, $row, $endColumn].Style.WrapText = $true
                $wsDashboard.Cells[$row, $startColumn, $row, $endColumn].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
            }

            if (($row % 2) -eq 1) {
                Set-CellFill -Range $wsDashboard.Cells[$row, 2, $row, 16] -HexColor '#F7F9FC'
            }

            $wsDashboard.Row($row).Height = 31
            $row++
        }
    }

    $wsDashboard.Column(1).Width = 2
    foreach ($column in 2..16) {
        $wsDashboard.Column($column).Width = 10
    }
    $wsDashboard.Column(2).Width = 14
    $wsDashboard.Column(16).Width = 12
    $wsDashboard.Cells['B39:P40'].Merge = $true
    $wsDashboard.Cells['B39'].Value = 'Klik op het tabblad “Actieve forwarding” voor alle details. Externe targets zijn rood gemarkeerd.'
    Set-CellFill -Range $wsDashboard.Cells['B39:P40'] -HexColor '#EAF2F8'
    $wsDashboard.Cells['B39:P40'].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#17365D'))
    $wsDashboard.Cells['B39:P40'].Style.Font.Italic = $true
    $wsDashboard.Cells['B39:P40'].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $wsDashboard.Cells['B39:P40'].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

    # Toelichting.
    $wsInfo = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName 'Toelichting'
    $wsInfo.View.ShowGridLines = $false
    $wsInfo.Cells['A1:H1'].Merge = $true
    $wsInfo.Cells['A1'].Value = 'TOELICHTING EN SCOPE'
    Set-CellFill -Range $wsInfo.Cells['A1:H1'] -HexColor '#17365D'
    $wsInfo.Cells['A1:H1'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $wsInfo.Cells['A1:H1'].Style.Font.Bold = $true
    $wsInfo.Cells['A1:H1'].Style.Font.Size = 20
    $wsInfo.Row(1).Height = 34

    $infoRows = @(
        @('Onderdeel', 'Beschrijving'),
        @('Mailbox-level forwarding', 'Controleert ForwardingAddress en ForwardingSmtpAddress, inclusief DeliverToMailboxAndForward.'),
        @('Inbox Rules', 'Controleert ForwardTo, RedirectTo en ForwardAsAttachmentTo. Verborgen regels worden meegenomen.'),
        @('Dashboardgrafieken', 'Worden als hoge-resolutie afbeeldingen ingebed om compatibiliteitsproblemen tussen ImportExcel- en EPPlus-versies te voorkomen.'),
        @('Intern / extern', 'Een target wordt intern geclassificeerd wanneer het domein voorkomt in de Exchange Accepted Domains.'),
        @('Risico-indicatie', 'Externe actieve forwarding krijgt de indicatie Verhoogd. Onherkenbare targets krijgen Controleren.'),
        @('Niet inbegrepen', 'Exchange transportregels, mailflow vanuit applicaties en forwarding buiten Exchange Online vallen buiten deze scan.'),
        @('Mogelijke scanbeperking', 'Get-InboxRule vereist passende Exchange RBAC-rechten. Eventuele fouten staan in het tabblad Scanlog.'),
        @('Rapport gegenereerd', (Get-Date -Format 'dd-MM-yyyy HH:mm:ss')),
        @('Mailboxen gescand', [string]$mailboxes.Count),
        @('Doorlooptijd', ("{0:n1} minuten" -f ($scanCompleted - $scanStarted).TotalMinutes))
    )

    $row = 3
    foreach ($infoRow in $infoRows) {
        $wsInfo.Cells[$row, 1, $row, 2].Merge = $true
        $wsInfo.Cells[$row, 3, $row, 8].Merge = $true
        $wsInfo.Cells[$row, 1].Value = $infoRow[0]
        $wsInfo.Cells[$row, 3].Value = $infoRow[1]
        $wsInfo.Cells[$row, 1, $row, 8].Style.WrapText = $true
        $wsInfo.Cells[$row, 1, $row, 8].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top

        if ($row -eq 3) {
            Set-CellFill -Range $wsInfo.Cells[$row, 1, $row, 8] -HexColor '#D9E1F2'
            $wsInfo.Cells[$row, 1, $row, 8].Style.Font.Bold = $true
        }
        elseif (($row % 2) -eq 0) {
            Set-CellFill -Range $wsInfo.Cells[$row, 1, $row, 8] -HexColor '#F7F9FC'
        }

        $wsInfo.Row($row).Height = if ($row -eq 3) { 24 } else { 40 }
        $row++
    }

    $wsInfo.Column(1).Width = 18
    $wsInfo.Column(2).Width = 18
    foreach ($column in 3..8) {
        $wsInfo.Column($column).Width = 15
    }

    # Workbook metadata en actieve tab.
    $excelPackage.Workbook.Properties.Title = 'Exchange Online Auto-Forwarding Audit'
    $excelPackage.Workbook.Properties.Subject = 'Tenantbreed rapport van mailbox-forwarding en Inbox Rules'
    $excelPackage.Workbook.Properties.Author = $env:USERNAME
    $excelPackage.Workbook.Properties.Company = 'Microsoft 365 Operations'
    $excelPackage.Workbook.Properties.Comments = 'Automatisch gegenereerd met Exchange Online PowerShell en ImportExcel.'
    $excelPackage.Workbook.Worksheets['Dashboard'].View.TabSelected = $true
    $excelPackage.Workbook.View.ActiveTab = 0

    Close-ExcelPackage -ExcelPackage $excelPackage

    Write-Host ''
    Write-Host 'Rapport succesvol gegenereerd:' -ForegroundColor Green
    Write-Host $OutputPath -ForegroundColor White
    Write-Host ("Actieve configuraties: {0} | Extern: {1} | Scanproblemen: {2}" -f $activeRecords.Count, $externalRecords.Count, $errorRecords.Count) -ForegroundColor Cyan

    if ($OpenReport) {
        Invoke-Item -LiteralPath $OutputPath
    }
}
catch {
    Write-Host ''
    Write-Host 'Het rapport kon niet worden gegenereerd.' -ForegroundColor Red
    Write-Host ("Fout: {0}" -f $_.Exception.Message) -ForegroundColor Red

    if ($_.InvocationInfo) {
        Write-Host ("Regel: {0}" -f $_.InvocationInfo.ScriptLineNumber) -ForegroundColor Yellow
        Write-Host ("Commando: {0}" -f $_.InvocationInfo.Line.Trim()) -ForegroundColor Yellow
    }

    throw
}
finally {
    foreach ($resource in @($dashboardImageResources)) {
        try {
            $resource.Dispose()
        }
        catch {
            # Opruimfouten mogen het rapportresultaat niet overschrijven.
        }
    }

    if ($dashboardAssetDirectory -and (Test-Path -LiteralPath $dashboardAssetDirectory)) {
        Remove-Item -LiteralPath $dashboardAssetDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }

    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
