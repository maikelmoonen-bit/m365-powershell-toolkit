#requires -Version 5.1
<#
.SYNOPSIS
    Entra Connect Upgrade - GUI Controller

.DESCRIPTION
    Interactieve WinForms-controller voor het begeleide upgradeproces van
    Microsoft Entra Connect (voorheen Azure AD Connect).

    Het script loopt een vast draaiboek van ruim 30 stappen door: voorbereiding,
    back-up van configuratie en sync-engine, scheduler en auto-upgrade beheren,
    de upgrade zelf, en de controles achteraf (delta/full sync, run history,
    deletion threshold). Handmatige stappen worden als checkpoint in de GUI
    getoond; geautomatiseerde stappen worden uitgevoerd en gevalideerd.

    Kenmerken:
      * Dry-run modus: elke stap wordt doorlopen zonder wijzigingen.
      * Starten vanaf een willekeurige stap (bijv. na een onderbreking).
      * ADSync-cmdlets draaien in een verborgen, niet-interactief childproces
        met time-out; de GUI blijft responsief.
      * Wacht automatisch tot de sync-engine idle is voordat scheduler-,
        sync- en configuratiestappen worden uitgevoerd, met retry bij
        'already running' / 'Configuration can not be changed'.
      * De installatiemap van de sync-engine wordt dynamisch gedetecteerd
        (registry -> ADSync-service ImagePath -> bekende kandidaatpaden).
      * Stopt het script ADSync en breekt de workflow af, dan wordt de service
        automatisch weer gestart.
      * Moderne UI met statusbadges, filter op aandachtspunten, detailvenster
        per stap, logvenster en HighDPI-ondersteuning.

.NOTES
    Vereisten: Windows PowerShell 5.1 of PowerShell 7, uit te voeren op de
    Entra Connect-server met lokale beheerrechten.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($env:OS -ne 'Windows_NT') {
    throw 'Dit script ondersteunt alleen Windows vanwege de WinForms-GUI.'
}

# ---------------------------------------------------------------------------
# STA-relaunch (vooral relevant voor PowerShell 7)
# ---------------------------------------------------------------------------
try {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        $exe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
        $relaunchArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', ('"{0}"' -f $PSCommandPath))
        Start-Process -FilePath $exe -ArgumentList ($relaunchArgs -join ' ')
        exit
    }
} catch {
    # Host ondersteunt geen relaunch-semantiek; WinForms werkt in PS 5.1 meestal alsnog.
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

try { [System.Windows.Forms.Application]::SetHighDpiMode([System.Windows.Forms.HighDpiMode]::SystemAware) | Out-Null } catch { }
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# ---------------------------------------------------------------------------
# Scriptstatus (vooraf initialiseren i.v.m. StrictMode)
# ---------------------------------------------------------------------------
$script:Results        = New-Object System.Collections.Generic.List[object]
$script:StartFromStep  = $null
$script:SyncPath       = $null
$script:Grid           = $null
$script:TxtLog         = $null
$script:TxtLocal       = $null
$script:TxtInstaller   = $null
$script:TxtCustomer    = $null
$script:TxtAadUpn      = $null
$script:ChkDelta       = $null
$script:ChkInitial     = $null
$script:ChkAutoUpgrade = $null
$script:ChkAdminCenter = $null
$script:ChkOnlyIssues  = $null
$script:CmbStartFrom   = $null
$script:Progress       = $null
$script:LblStatus      = $null
$script:LblSyncPath    = $null
$script:ActionControls = @()
$script:ServiceStoppedByScript = $false
$script:PurgeSetEarly  = $false

# ---------------------------------------------------------------------------
# Designtokens
# ---------------------------------------------------------------------------
$script:Clr = @{
    Bg          = [System.Drawing.Color]::FromArgb(243, 244, 246)
    Card        = [System.Drawing.Color]::White
    Border      = [System.Drawing.Color]::FromArgb(225, 227, 231)
    Text        = [System.Drawing.Color]::FromArgb(32, 31, 30)
    Muted       = [System.Drawing.Color]::FromArgb(110, 112, 118)
    Accent      = [System.Drawing.Color]::FromArgb(15, 108, 189)
    AccentDark  = [System.Drawing.Color]::FromArgb(11, 84, 148)
    AccentLight = [System.Drawing.Color]::FromArgb(235, 243, 252)
    Danger      = [System.Drawing.Color]::FromArgb(196, 43, 28)
    DangerDark  = [System.Drawing.Color]::FromArgb(163, 33, 20)
    Header      = [System.Drawing.Color]::FromArgb(23, 42, 69)
    LogBg       = [System.Drawing.Color]::FromArgb(28, 30, 34)
    LogText     = [System.Drawing.Color]::FromArgb(214, 219, 226)
    OkBg        = [System.Drawing.Color]::FromArgb(223, 246, 221)
    OkFg        = [System.Drawing.Color]::FromArgb(15, 92, 43)
    WarnBg      = [System.Drawing.Color]::FromArgb(255, 244, 206)
    WarnFg      = [System.Drawing.Color]::FromArgb(129, 87, 4)
    BlockBg     = [System.Drawing.Color]::FromArgb(253, 231, 229)
    BlockFg     = [System.Drawing.Color]::FromArgb(154, 32, 20)
    ManualBg    = [System.Drawing.Color]::FromArgb(230, 241, 252)
    ManualFg    = [System.Drawing.Color]::FromArgb(16, 78, 139)
    InfoBg      = [System.Drawing.Color]::FromArgb(241, 242, 244)
    InfoFg      = [System.Drawing.Color]::FromArgb(90, 92, 98)
    RowAlt      = [System.Drawing.Color]::FromArgb(250, 250, 251)
}

$script:Fonts = @{
    Base    = New-Object System.Drawing.Font('Segoe UI', 9.5)
    Bold    = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
    Title   = New-Object System.Drawing.Font('Segoe UI Light', 17)
    Sub     = New-Object System.Drawing.Font('Segoe UI', 9)
    Caption = New-Object System.Drawing.Font('Segoe UI', 8.25)
    Mono    = New-Object System.Drawing.Font('Consolas', 9.5)
    Section = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
}

# ---------------------------------------------------------------------------
# Stap-mapping
# ---------------------------------------------------------------------------
$script:StepMap = [ordered]@{
    'Stap 1  - Documentatiemap voorbereiden'           = 'Stap 1'
    'Stap 2  - Monitoring in onderhoudsmodus'          = 'Stap 2'
    'Stap 2a - Tijdelijke rechten toekennen'           = 'Stap 2a'
    'Stap 3  - RDP / juiste server bevestigen'         = 'Stap 3'
    'Stap 4  - ADSyncAdmins controleren'               = 'Stap 4'
    'Stap 5  - Installer downloaden/valideren'         = 'Stap 5'
    'Stap 6  - Vrije ruimte C-schijf'                  = 'Stap 6'
    'Stap 6b - Staging mode inschakelen'               = 'Stap 6b'
    'Stap 7  - PowerShell als Administrator'           = 'Stap 7'
    'Stap 8  - Scheduler uitzetten'                    = 'Stap 8'
    'Stap 9  - ADSync service stoppen'                 = 'Stap 9'
    'Stap 10 - AADConnect folder zippen'               = 'Stap 10'
    'Stap 11 - Sync-installatiemap zippen (optioneel)' = 'Stap 11'
    'Stap 12 - ADSync service starten'                 = 'Stap 12'
    'Stap 13-15 - Configuratie exporteren (.json)'     = 'Stap 13-15'
    'Stap 15 - JSON en trace-log veiligstellen'        = 'Stap 15'
    'Stap 16 - Deletion threshold controleren'         = 'Stap 16'
    'Stap 17 - Connector export (.csv)'                = 'Stap 17'
    'Stap 18-19 - Licentiegroep-exports (M365)'        = 'Stap 18-19'
    'Stap 20-25 - Installer starten / wizard'          = 'Stap 20-25'
    'Stap 26 - Scheduler weer inschakelen'             = 'Stap 26'
    'Stap 27 - Auto Upgrade uitzetten'                 = 'Stap 27'
    'Stap 28 - Monitoring uit onderhoudsmodus'         = 'Stap 28'
    'Stap 29 - Delta sync starten'                     = 'Stap 29'
    'Stap 30 - Initial sync starten'                   = 'Stap 30'
    'Stap 31 - Excelsheet bijwerken'                   = 'Stap 31'
    'Stap 32 - Teammail versturen'                     = 'Stap 32'
    'Stap 33 - Run history op 30 dagen'                = 'Stap 33'
    'Stap 34 - Tijdelijke rechten opruimen'            = 'Stap 34'
}

# ---------------------------------------------------------------------------
# Helpers: status en stapvolgorde
# ---------------------------------------------------------------------------
function Normalize-Status {
    param([string]$Status)
    $statusValue = if ($null -eq $Status) { '' } else { [string]$Status }
    switch -Regex ($statusValue.Trim().ToUpperInvariant()) {
        '^OK$'              { 'OK'; break }
        '^LET(\s|_|-)?OP$'  { 'LET OP'; break }
        '^LET$'             { 'LET OP'; break }
        '^WARNING$'         { 'LET OP'; break }
        '^BLOKKADE$'        { 'BLOKKADE'; break }
        '^HANDMATIG$'       { 'HANDMATIG'; break }
        '^INFO$'            { 'INFO'; break }
        default             { 'INFO' }
    }
}

function Get-StatusColors {
    param([string]$Status)
    switch ($Status) {
        'OK'        { return @{ Bg = $script:Clr.OkBg;     Fg = $script:Clr.OkFg } }
        'LET OP'    { return @{ Bg = $script:Clr.WarnBg;   Fg = $script:Clr.WarnFg } }
        'BLOKKADE'  { return @{ Bg = $script:Clr.BlockBg;  Fg = $script:Clr.BlockFg } }
        'HANDMATIG' { return @{ Bg = $script:Clr.ManualBg; Fg = $script:Clr.ManualFg } }
        default     { return @{ Bg = $script:Clr.InfoBg;   Fg = $script:Clr.InfoFg } }
    }
}

function Test-StepActive {
    param([string]$StepId)
    if ([string]::IsNullOrWhiteSpace($script:StartFromStep)) { return $true }

    $keys = @($script:StepMap.Keys)
    $startKey = ($keys | Where-Object { $script:StepMap[$_] -eq $script:StartFromStep }) | Select-Object -First 1
    $thisKey  = ($keys | Where-Object { $script:StepMap[$_] -eq $StepId })              | Select-Object -First 1

    if ([string]::IsNullOrWhiteSpace($startKey) -or [string]::IsNullOrWhiteSpace($thisKey)) { return $true }

    $startIdx = [array]::IndexOf($keys, $startKey)
    $thisIdx  = [array]::IndexOf($keys, $thisKey)
    return ($thisIdx -ge $startIdx)
}

function Get-StepProgressValue {
    param([string]$StepId)
    $keys = @($script:StepMap.Keys)
    $thisKey = ($keys | Where-Object { $script:StepMap[$_] -eq $StepId }) | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($thisKey)) { return $null }
    return ([array]::IndexOf($keys, $thisKey) + 1)
}

# ---------------------------------------------------------------------------
# Logging en resultaten
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Ok', 'Warn', 'Error', 'Muted')]
        [string]$Level = 'Info'
    )
    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $line = "[$timestamp] $Message"
    if ($script:TxtLog) {
        $color = switch ($Level) {
            'Ok'    { [System.Drawing.Color]::FromArgb(126, 211, 143) }
            'Warn'  { [System.Drawing.Color]::FromArgb(242, 201, 106) }
            'Error' { [System.Drawing.Color]::FromArgb(240, 129, 118) }
            'Muted' { [System.Drawing.Color]::FromArgb(140, 145, 155) }
            default { $script:Clr.LogText }
        }
        $script:TxtLog.SelectionStart  = $script:TxtLog.TextLength
        $script:TxtLog.SelectionLength = 0
        $script:TxtLog.SelectionColor  = $color
        $script:TxtLog.AppendText($line + [Environment]::NewLine)
        $script:TxtLog.SelectionStart  = $script:TxtLog.TextLength
        $script:TxtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Set-StatusText {
    param([string]$Text)
    if ($script:LblStatus) {
        $script:LblStatus.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Test-RowVisible {
    param([string]$Status)
    if ($script:ChkOnlyIssues -and $script:ChkOnlyIssues.Checked) {
        return ($Status -in @('LET OP', 'BLOKKADE'))
    }
    return $true
}

function Add-GridRow {
    param($ResultObject)
    if (-not $script:Grid) { return }
    if (-not (Test-RowVisible -Status $ResultObject.Status)) { return }

    # Alleen meescrollen als de gebruiker al onderaan stond. Wie terugscrolt
    # naar eerdere stappen, houdt zijn positie.
    $stickToBottom = $true
    $rowCount = $script:Grid.Rows.Count
    if ($rowCount -gt 0) {
        $first = $script:Grid.FirstDisplayedScrollingRowIndex
        if ($first -ge 0) {
            $lastShown = $first + $script:Grid.DisplayedRowCount($true) - 1
            $stickToBottom = ($lastShown -ge ($rowCount - 1))
        }
    }

    $idx = $script:Grid.Rows.Add(
        $ResultObject.Timestamp.ToString('HH:mm:ss'),
        $ResultObject.Step,
        $ResultObject.Title,
        $ResultObject.Status,
        ($ResultObject.Detail -replace '\r?\n', ' ')
    )
    $row = $script:Grid.Rows[$idx]
    $row.Tag = $ResultObject

    $colors = Get-StatusColors -Status $ResultObject.Status
    $statusCell = $row.Cells[3]
    $statusCell.Style.BackColor          = $colors.Bg
    $statusCell.Style.ForeColor          = $colors.Fg
    $statusCell.Style.SelectionBackColor = $colors.Bg
    $statusCell.Style.SelectionForeColor = $colors.Fg
    $statusCell.Style.Font               = $script:Fonts.Bold
    $statusCell.Style.Alignment          = 'MiddleCenter'

    if ($ResultObject.Status -eq 'BLOKKADE') {
        $row.Cells[2].Style.ForeColor = $script:Clr.BlockFg
        $row.Cells[2].Style.Font      = $script:Fonts.Bold
    }

    if ($stickToBottom) {
        try { $script:Grid.FirstDisplayedScrollingRowIndex = $script:Grid.Rows.Count - 1 } catch { }
    }
}

function Update-GridFilter {
    if (-not $script:Grid) { return }
    $script:Grid.Rows.Clear()
    foreach ($r in $script:Results) { Add-GridRow -ResultObject $r }
    if ($script:Grid.Rows.Count -gt 0) {
        try { $script:Grid.FirstDisplayedScrollingRowIndex = $script:Grid.Rows.Count - 1 } catch { }
        $script:Grid.ClearSelection()
    }
}

function Set-UiBusy {
    param([bool]$Busy)
    foreach ($c in @($script:ActionControls)) {
        if ($c) { $c.Enabled = -not $Busy }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function New-Result {
    param(
        [string]$Step,
        [string]$Title,
        [string]$Status,
        [string]$Detail
    )

    $s = Normalize-Status $Status
    $obj = [pscustomobject]@{
        Timestamp = (Get-Date)
        Step      = $Step
        Title     = $Title
        Status    = $s
        Detail    = $Detail
    }
    $script:Results.Add($obj) | Out-Null
    Add-GridRow -ResultObject $obj

    $level = switch ($s) {
        'OK'        { 'Ok' }
        'LET OP'    { 'Warn' }
        'BLOKKADE'  { 'Error' }
        'HANDMATIG' { 'Info' }
        default     { 'Muted' }
    }
    Write-Log ("{0,-9} | {1} - {2}" -f $s, $Step, $Title) -Level $level
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        Write-Log ("            {0}" -f ($Detail -replace '\r?\n', ' ')) -Level 'Muted'
    }

    $p = Get-StepProgressValue -StepId $Step
    if ($null -ne $p -and $script:Progress) {
        $script:Progress.Value = [Math]::Min($script:Progress.Maximum, $p)
    }
    Set-StatusText ("$Step - $Title")
}

function Clear-Results {
    $script:Results.Clear()
    if ($script:Grid)     { $script:Grid.Rows.Clear() }
    if ($script:TxtLog)   { $script:TxtLog.Clear() }
    if ($script:Progress) { $script:Progress.Value = 0 }
    Set-StatusText 'Gereed.'
}

function Write-Summary {
    $summary = @($script:Results | Group-Object Status | Sort-Object Name)
    Write-Log ('-' * 78) -Level 'Muted'
    Write-Log 'SAMENVATTING'
    foreach ($group in $summary) {
        $lvl = switch ($group.Name) {
            'OK'       { 'Ok' }
            'LET OP'   { 'Warn' }
            'BLOKKADE' { 'Error' }
            default    { 'Muted' }
        }
        Write-Log ('  {0,-10}: {1}' -f $group.Name, $group.Count) -Level $lvl
    }

    $blocked = @($script:Results | Where-Object { $_.Status -eq 'BLOKKADE' }).Count
    if ($blocked -gt 0) {
        Write-Log 'Eindoordeel: de echte run is nu niet verstandig. Los eerst de blokkades op.' -Level 'Error'
    } else {
        Write-Log 'Eindoordeel: er zijn geen harde blokkades zichtbaar. Handmatige stappen blijven vereist.' -Level 'Ok'
    }
}

# ---------------------------------------------------------------------------
# Omgevingsdetectie
# ---------------------------------------------------------------------------
function Show-StepDialog {
    param(
        [string]$Title,
        [string]$Text,
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    return [System.Windows.Forms.MessageBox]::Show($Text, $Title, $Buttons, $Icon)
}

function Get-InputState {
    [pscustomobject]@{
        LocalExportFolder  = $script:TxtLocal.Text.Trim()
        InstallerPath      = $script:TxtInstaller.Text.Trim()
        CustomerName       = $script:TxtCustomer.Text.Trim()
        AadUpn             = $script:TxtAadUpn.Text.Trim()
        StartDelta         = $script:ChkDelta.Checked
        StartInitial       = $script:ChkInitial.Checked
        DisableAutoUpgrade = $script:ChkAutoUpgrade.Checked
        OpenAdminCenter    = $script:ChkAdminCenter.Checked
    }
}

function Test-IsAdministrator {
    try {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    } catch {
        return $false
    }
}

function Get-FreeSpaceGB {
    try {
        $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
        if ($null -eq $drive) { return $null }
        return [math]::Round(($drive.FreeSpace / 1GB), 1)
    } catch {
        return $null
    }
}

function Resolve-SyncRootFolder {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = $Path.Trim().Trim('"')
    try {
        if (Test-Path -LiteralPath $p -PathType Leaf) { $p = Split-Path -LiteralPath $p -Parent }
    } catch { return $null }
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    # miiserver.exe staat in <root>\Bin -> een niveau omhoog
    $leaf = Split-Path -LiteralPath $p -Leaf
    if ($leaf -in @('Bin', 'bin', 'SyncBin')) {
        $parent = Split-Path -LiteralPath $p -Parent
        if ($parent -and (Test-Path -LiteralPath $parent)) { $p = $parent }
    }
    return $p
}

<#
    Bepaalt de installatiemap van de sync-engine.
    Nieuwere Entra Connect-versies gebruiken niet meer per definitie
    'C:\Program Files\Microsoft Azure AD Sync'. Daarom detecteren we het pad
    dynamisch en behandelen we een ontbrekend pad als optioneel, niet fataal.
#>
function Get-EntraConnectSyncPath {
    # 1. Registry (product-installatiesleutels)
    $regKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Azure AD Connect',
        'HKLM:\SOFTWARE\Microsoft\Microsoft Entra Connect',
        'HKLM:\SOFTWARE\Microsoft\Azure AD Sync',
        'HKLM:\SOFTWARE\Microsoft\Microsoft Entra Connect Sync'
    )
    $regValues = @('InstallPath', 'SyncPath', 'WizardPath', 'Location', 'InstallDir')

    foreach ($key in $regKeys) {
        try {
            if (-not (Test-Path -LiteralPath $key)) { continue }
            $item = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
            if ($null -eq $item) { continue }
            $names = @($item.PSObject.Properties.Name)
            foreach ($valueName in $regValues) {
                if ($names -notcontains $valueName) { continue }
                $resolved = Resolve-SyncRootFolder ([string]$item.$valueName)
                if ($resolved) { return $resolved }
            }
        } catch { }
    }

    # 2. ImagePath van de ADSync-service
    try {
        $svcKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\ADSync'
        if (Test-Path -LiteralPath $svcKey) {
            $svcItem = Get-ItemProperty -LiteralPath $svcKey -ErrorAction SilentlyContinue
            if ($svcItem -and (@($svcItem.PSObject.Properties.Name) -contains 'ImagePath')) {
                $image = [string]$svcItem.ImagePath
                # Strip eventuele argumenten en aanhalingstekens
                if ($image -match '^\s*"([^"]+)"') { $image = $Matches[1] }
                elseif ($image -match '^\s*(\S+\.exe)') { $image = $Matches[1] }
                $resolved = Resolve-SyncRootFolder $image
                if ($resolved) { return $resolved }
            }
        }
    } catch { }

    # 3. Bekende kandidaatpaden (oud -> nieuw)
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Microsoft Azure AD Sync'),
        (Join-Path $env:ProgramFiles 'Microsoft Azure AD Connect'),
        (Join-Path $env:ProgramFiles 'Microsoft Azure AD Connect Sync'),
        (Join-Path $env:ProgramFiles 'Microsoft Entra\Connect Sync'),
        (Join-Path $env:ProgramFiles 'Microsoft Entra Connect'),
        (Join-Path $env:ProgramFiles 'Microsoft Entra Connect Sync')
    )
    foreach ($c in $candidates) {
        try { if (Test-Path -LiteralPath $c) { return $c } } catch { }
    }

    return $null
}

function Update-SyncPathLabel {
    $script:SyncPath = Get-EntraConnectSyncPath
    if ($script:LblSyncPath) {
        if ($script:SyncPath) {
            $script:LblSyncPath.Text      = "Sync-installatiemap gedetecteerd: $($script:SyncPath)"
            $script:LblSyncPath.ForeColor = $script:Clr.OkFg
        } else {
            $script:LblSyncPath.Text      = 'Sync-installatiemap: niet gevonden (normaal bij nieuwere Entra Connect-versies) - stap 11 wordt overgeslagen.'
            $script:LblSyncPath.ForeColor = $script:Clr.Muted
        }
    }
    return $script:SyncPath
}

function Test-ADSyncModuleAvailable {
    try { return [bool](Get-Module -ListAvailable -Name ADSync) } catch { return $false }
}

function Get-ADSyncServiceSafe {
    try { return Get-Service -Name ADSync -ErrorAction Stop } catch { return $null }
}

function Get-WindowsPowerShellPath {
    $candidate = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $candidate) { return $candidate }
    return 'powershell.exe'
}

<#
    Voert ADSync-cmdlets uit in een apart Windows PowerShell-proces.

    Het childproces draait zonder eigen consolevenster en met
    -NonInteractive. Ontbreekt een verplichte parameter, dan faalt de aanroep
    direct met een duidelijke fout in plaats van stil te wachten op invoer in
    het consolevenster achter de GUI. Output gaat via een tijdelijk bestand
    terug naar de GUI. Tijdens het wachten blijft de GUI responsief en er
    geldt een time-out. Grafische vensters van het childproces (zoals een
    Microsoft-aanmeldvenster) worden gewoon getoond.
#>
function Invoke-ADSyncChild {
    param(
        [string]$ScriptText,
        [int]$TimeoutSeconds = 300,
        [string]$WaitMessage
    )

    if (-not (Test-ADSyncModuleAvailable)) {
        return [pscustomobject]@{ Success = $false; ExitCode = 1; TimedOut = $false; Output = 'ADSync module niet gevonden.' }
    }

    $id      = [guid]::NewGuid().Guid
    $temp    = Join-Path $env:TEMP ("ADSyncChild-$id.ps1")
    $outFile = Join-Path $env:TEMP ("ADSyncChild-$id.out")
    $safeOut = $outFile.Replace("'", "''")

    $wrapper = @"
`$ErrorActionPreference = 'Stop'
`$ProgressPreference    = 'SilentlyContinue'
`$__outFile = '$safeOut'
try {
    Import-Module ADSync -ErrorAction Stop
    & {
$ScriptText
    } *>&1 | Out-String -Width 4096 | Set-Content -LiteralPath `$__outFile -Encoding UTF8
    exit 0
} catch {
    (`$_ | Out-String -Width 4096) | Add-Content -LiteralPath `$__outFile -Encoding UTF8
    exit 1
}
"@

    Set-Content -LiteralPath $temp -Value $wrapper -Encoding UTF8
    $proc = $null
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName        = Get-WindowsPowerShellPath
        $psi.Arguments       = ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $temp)
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow  = $true

        $proc     = [System.Diagnostics.Process]::Start($psi)
        $sw       = [System.Diagnostics.Stopwatch]::StartNew()
        $timedOut = $false
        $lastTick = -1

        while (-not $proc.WaitForExit(200)) {
            [System.Windows.Forms.Application]::DoEvents()
            $elapsed = [int]$sw.Elapsed.TotalSeconds
            if ($WaitMessage -and $elapsed -ne $lastTick) {
                $lastTick = $elapsed
                Set-StatusText ("{0}  ({1}s)" -f $WaitMessage, $elapsed)
            }
            if ($TimeoutSeconds -gt 0 -and $sw.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
                try { $proc.Kill() } catch { }
                $timedOut = $true
                break
            }
        }
        if (-not $timedOut) { $proc.WaitForExit() }

        $output = ''
        if (Test-Path -LiteralPath $outFile) {
            $output = (Get-Content -LiteralPath $outFile -Raw -ErrorAction SilentlyContinue)
            if ($null -eq $output) { $output = '' }
        }
        $output = $output.Trim()

        if ($timedOut) {
            return [pscustomobject]@{
                Success  = $false
                ExitCode = -1
                TimedOut = $true
                Output   = "Time-out na $TimeoutSeconds seconden; het childproces is beeindigd. $output".Trim()
            }
        }

        return [pscustomobject]@{
            Success  = ($proc.ExitCode -eq 0)
            ExitCode = $proc.ExitCode
            TimedOut = $false
            Output   = $output
        }
    } finally {
        if ($proc) { $proc.Dispose() }
        Remove-Item -LiteralPath $temp    -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    }
}

<#
    Vraagt (of bevestigt) het Entra ID-account voor stap 16 in een eigen,
    altijd-bovenop venster. Geeft de UPN terug, of $null bij Overslaan.
#>
function Show-AadUpnPrompt {
    param([string]$Default)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Stap 16 - Deletion threshold uitlezen'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.StartPosition   = 'CenterScreen'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false
    $dlg.TopMost         = $true
    $dlg.Font            = $script:Fonts.Base
    $dlg.BackColor       = $script:Clr.Card
    $dlg.ClientSize      = New-Object System.Drawing.Size(500, 214)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text      = "Get-ADSyncExportDeletionThreshold vereist een Entra ID-account (UPN) met voldoende rechten, bijvoorbeeld het tijdelijke beheeraccount uit stap 2a.`r`n`r`nNa 'Doorgaan' verschijnt mogelijk een Microsoft-aanmeldvenster (eventueel met MFA). Zie je het niet, kijk dan in de taakbalk."
    $lblInfo.Location  = New-Object System.Drawing.Point(16, 14)
    $lblInfo.Size      = New-Object System.Drawing.Size(468, 84)
    $lblInfo.ForeColor = $script:Clr.Text
    $dlg.Controls.Add($lblInfo)

    $lblUpn = New-Object System.Windows.Forms.Label
    $lblUpn.Text      = 'ENTRA-ACCOUNT (UPN)'
    $lblUpn.Font      = $script:Fonts.Caption
    $lblUpn.ForeColor = $script:Clr.Muted
    $lblUpn.Location  = New-Object System.Drawing.Point(16, 104)
    $lblUpn.AutoSize  = $true
    $dlg.Controls.Add($lblUpn)

    $txtUpn = New-Object System.Windows.Forms.TextBox
    $txtUpn.Location    = New-Object System.Drawing.Point(16, 122)
    $txtUpn.Width       = 468
    $txtUpn.BorderStyle = 'FixedSingle'
    $txtUpn.Text        = $Default
    $dlg.Controls.Add($txtUpn)

    $lblErr = New-Object System.Windows.Forms.Label
    $lblErr.Text      = ''
    $lblErr.Font      = $script:Fonts.Caption
    $lblErr.ForeColor = $script:Clr.Danger
    $lblErr.Location  = New-Object System.Drawing.Point(16, 150)
    $lblErr.AutoSize  = $true
    $dlg.Controls.Add($lblErr)

    $btnOk   = New-ModernButton -Text 'Doorgaan'  -X 272 -Y 172 -W 102 -H 30 -Kind 'Primary'
    $btnSkip = New-ModernButton -Text 'Overslaan' -X 382 -Y 172 -W 102 -H 30
    $btnSkip.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnOk)
    $dlg.Controls.Add($btnSkip)
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnSkip

    $btnOk.Add_Click({
        if ($txtUpn.Text.Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        } else {
            $lblErr.Text = 'Vul een geldige UPN in, bijvoorbeeld naam@domein.nl.'
            $txtUpn.Focus() | Out-Null
        }
    })
    $dlg.Add_Shown({ $dlg.Activate(); $txtUpn.Focus() | Out-Null; $txtUpn.SelectAll() })

    try {
        $result = $dlg.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $txtUpn.Text.Trim() }
        return $null
    } finally {
        $dlg.Dispose()
    }
}

<#
    Maakt een leesbare samenvatting van een fout uit een ADSync-childproces.
#>
function Get-ADSyncErrorSummary {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return 'Onbekende fout (geen uitvoer van het childproces).' }
    $clean = $Raw -replace '</?error>', '' -replace 'System\.Management\.Automation\.CmdletInvocationException:\s*', ''
    $first = (($clean -split '\r?\n') | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -First 2) -join ' '
    if ($first.Length -gt 400) { $first = $first.Substring(0, 400) + '...' }
    if ($Raw -match 'Configuration can not be changed|Cannot change configuration') {
        return "Er liep op dat moment een synchronisatiecyclus. Zolang een sync loopt, kan de configuratie niet worden gewijzigd. Origineel: $first"
    }
    if ($Raw -match 'already running|is busy|SyncCycleInProgress') {
        return "Er loopt al een synchronisatie. Origineel: $first"
    }
    return $first
}

<#
    Wacht een aantal seconden zonder de GUI te bevriezen.
#>
function Wait-Responsive {
    param([int]$Seconds, [string]$Message)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastTick = -1
    while ($sw.Elapsed.TotalSeconds -lt $Seconds) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
        if ($Message) {
            $t = [int]$sw.Elapsed.TotalSeconds
            if ($t -ne $lastTick) { $lastTick = $t; Set-StatusText $Message }
        }
    }
}

<#
    Bepaalt of de sync-engine bezig is. Leeg resultaat van
    Get-ADSyncConnectorRunStatus = idle (zie Microsoft-docs over de scheduler).
    Daarnaast wordt SyncCycleInProgress van Get-ADSyncScheduler meegenomen
    (indien die eigenschap bestaat), zodat ook de pauzes tussen connector-runs
    binnen een cyclus als 'bezig' tellen.
    Geeft @{ Busy; Names; Error } terug.
#>
function Get-ADSyncBusyState {
    $probe = @'
$running = @(Get-ADSyncConnectorRunStatus | Where-Object { $null -ne $_ })
$names = @($running | ForEach-Object {
    if ($_.PSObject.Properties['ConnectorName']) { [string]$_.ConnectorName } else { [string]$_ }
})
$inProgress = $false
try {
    $sch = Get-ADSyncScheduler
    if ($sch -and $sch.PSObject.Properties['SyncCycleInProgress']) { $inProgress = [bool]$sch.SyncCycleInProgress }
} catch { }
if ($running.Count -gt 0 -or $inProgress) {
    Write-Output ('ADSYNCSTATE=BUSY;' + ($names -join ', '))
} else {
    Write-Output 'ADSYNCSTATE=IDLE;'
}
'@
    $r = Invoke-ADSyncChild -ScriptText $probe -TimeoutSeconds 120
    if ($r.Output -match '(?m)^ADSYNCSTATE=(BUSY|IDLE);(.*)$') {
        $names = $Matches[2].Trim()
        if ([string]::IsNullOrWhiteSpace($names)) { $names = 'scheduler-cyclus' }
        return [pscustomobject]@{ Busy = ($Matches[1] -eq 'BUSY'); Names = $names; Error = $null }
    }
    return [pscustomobject]@{ Busy = $false; Names = ''; Error = (Get-ADSyncErrorSummary $r.Output) }
}

<#
    Wacht tot de sync-engine idle is. De GUI blijft bruikbaar. Na elk blok van
    $ChunkMinutes vraagt het script of je langer wilt wachten.
    Retourneert $true (idle, verder gaan) of $false (gebruiker kiest overslaan).
#>
function Wait-ADSyncIdle {
    param(
        [string]$StepLabel,
        [int]$ChunkMinutes = 10,
        [int]$PollSeconds  = 15
    )

    if (-not (Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncConnectorRunStatus')) {
        Write-Log "${StepLabel}: Get-ADSyncConnectorRunStatus niet beschikbaar; kan niet controleren of er een sync loopt." -Level 'Warn'
        return $true
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $announced = $false
    $nextPromptAt = $ChunkMinutes * 60

    while ($true) {
        $state = Get-ADSyncBusyState
        if ($state.Error) {
            Write-Log "${StepLabel}: sync-status kon niet worden bepaald ($($state.Error)). Er wordt toch doorgegaan." -Level 'Warn'
            return $true
        }
        if (-not $state.Busy) {
            if ($announced) {
                Write-Log ("{0}: lopende synchronisatie is klaar na {1:N0} min. Er wordt doorgegaan." -f $StepLabel, $sw.Elapsed.TotalMinutes) -Level 'Ok'
            }
            return $true
        }

        if (-not $announced) {
            Write-Log "${StepLabel}: er loopt een synchronisatie ($($state.Names)). Wachten tot deze klaar is..." -Level 'Warn'
            $announced = $true
        }

        if ($sw.Elapsed.TotalSeconds -ge $nextPromptAt) {
            $msg = ("Er loopt na {0:N0} minuten nog steeds een synchronisatie ({1}).`r`n`r`nJa = nog {2} minuten wachten`r`nNee = deze stap overslaan (later handmatig uitvoeren)" -f $sw.Elapsed.TotalMinutes, $state.Names, $ChunkMinutes)
            $answer = Show-StepDialog -Title "$StepLabel - sync loopt nog" -Text $msg -Buttons YesNo -Icon Question
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
            $nextPromptAt = $sw.Elapsed.TotalSeconds + ($ChunkMinutes * 60)
        }

        $elapsedMin = [int]$sw.Elapsed.TotalMinutes
        Wait-Responsive -Seconds $PollSeconds -Message ("{0} - wacht tot lopende sync klaar is ({1}), {2} min" -f $StepLabel, $state.Names, $elapsedMin)
    }
}

function Get-ADSyncPurgeIntervalDays {
    $r = Invoke-ADSyncChild -ScriptText 'Write-Output ("PURGE=" + ([TimeSpan](Get-ADSyncScheduler).PurgeRunHistoryInterval).TotalDays)' -TimeoutSeconds 120
    if ($r.Output -match '(?m)^PURGE=([\d\.,]+)') {
        return [double]::Parse(($Matches[1] -replace ',', '.'), [Globalization.CultureInfo]::InvariantCulture)
    }
    return $null
}

<#
    Zet PurgeRunHistoryInterval op 30 dagen. Wacht eerst tot er geen sync
    loopt en probeert het opnieuw als de engine alsnog 'Configuration can not
    be changed' meldt (er is dan net een cyclus gestart).
#>
function Set-ADSyncRunHistory30 {
    param([string]$StepLabel, [int]$MaxAttempts = 3)

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        if (-not (Wait-ADSyncIdle -StepLabel $StepLabel)) {
            return [pscustomobject]@{ Success = $false; Skipped = $true; Detail = 'Overgeslagen: er liep nog een synchronisatie. Voer later uit: Set-ADSyncScheduler -PurgeRunHistoryInterval 30.00:00:00' }
        }
        $r = Invoke-ADSyncChild -ScriptText 'Set-ADSyncScheduler -PurgeRunHistoryInterval 30.00:00:00' -WaitMessage "$StepLabel - run history instellen"
        if ($r.Success) {
            $days = Get-ADSyncPurgeIntervalDays
            if ($null -ne $days -and [math]::Abs($days - 30) -lt 0.01) {
                return [pscustomobject]@{ Success = $true; Skipped = $false; Detail = 'PurgeRunHistoryInterval staat op 30 dagen (gecontroleerd met Get-ADSyncScheduler).' }
            }
            return [pscustomobject]@{ Success = $true; Skipped = $false; Detail = "Set-ADSyncScheduler uitgevoerd; controlewaarde: $days dagen." }
        }
        $isBusyError = ($r.Output -match 'Configuration can not be changed|Cannot change configuration|already running')
        if (-not $isBusyError -or $attempt -eq $MaxAttempts) {
            return [pscustomobject]@{ Success = $false; Skipped = $false; Detail = (Get-ADSyncErrorSummary $r.Output) }
        }
        Write-Log "${StepLabel}: er startte net een sync-cyclus; poging $attempt van $MaxAttempts mislukt, opnieuw wachten..." -Level 'Warn'
        Wait-Responsive -Seconds 10 -Message "$StepLabel - korte pauze voor nieuwe poging"
    }
}

<#
    Start een sync-cyclus zodra de engine idle is. Na het inschakelen van de
    scheduler (stap 26) start die zelf vrijwel direct een cyclus; daarom wordt
    eerst gewacht en bij 'already running' nog een keer opnieuw geprobeerd.
    Retourneert @{ Status; Detail } voor New-Result.
#>
function Start-ADSyncCycleWhenIdle {
    param(
        [ValidateSet('Delta', 'Initial')][string]$PolicyType,
        [string]$StepLabel,
        [int]$MaxAttempts = 3
    )

    # Even de scheduler de kans geven een net gestarte cyclus te registreren.
    Wait-Responsive -Seconds 10 -Message "$StepLabel - sync-status controleren"

    $lastError = ''
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        if (-not (Wait-ADSyncIdle -StepLabel $StepLabel)) {
            return [pscustomobject]@{ Status = 'LET OP'; Detail = "Overgeslagen door gebruiker; er liep nog een synchronisatie. Start later handmatig: Start-ADSyncSyncCycle -PolicyType $PolicyType" }
        }
        $r = Invoke-ADSyncChild -ScriptText "Start-ADSyncSyncCycle -PolicyType $PolicyType | Out-String -Width 4096" -WaitMessage "$StepLabel - $PolicyType sync starten"
        if ($r.Success) {
            $extra = (($r.Output -replace '\s+', ' ').Trim())
            return [pscustomobject]@{ Status = 'OK'; Detail = ("$PolicyType sync gestart. $extra").Trim() }
        }
        $lastError = Get-ADSyncErrorSummary $r.Output
        if ($r.Output -notmatch 'already running|is busy|in progress|Configuration can not be changed') { break }
        Write-Log "${StepLabel}: er startte net een andere cyclus; poging $attempt van $MaxAttempts, opnieuw wachten..." -Level 'Warn'
        Wait-Responsive -Seconds 15 -Message "$StepLabel - korte pauze voor nieuwe poging"
    }
    return [pscustomobject]@{ Status = 'LET OP'; Detail = $lastError }
}

function Restore-ADSyncServiceIfNeeded {
    if (-not $script:ServiceStoppedByScript) { return }
    try {
        $svc = Get-ADSyncServiceSafe
        if ($svc -and $svc.Status -ne 'Running') {
            Write-Log 'Workflow is afgebroken terwijl ADSync door het script was gestopt - service wordt opnieuw gestart.' -Level 'Warn'
            Start-Service -Name ADSync -ErrorAction Stop
            New-Result 'Herstel' 'ADSync service herstart na afbreken' 'LET OP' 'De service is opnieuw gestart. De sync-scheduler staat nog UIT (stap 8). Hervat via "Start vanaf stap" of zet hem bewust aan met Set-ADSyncScheduler -SyncCycleEnabled $true.'
        }
    } catch {
        New-Result 'Herstel' 'ADSync service herstarten' 'BLOKKADE' "ADSync kon niet automatisch worden herstart: $($_.Exception.Message)"
    } finally {
        $script:ServiceStoppedByScript = $false
    }
}

function Test-ADSyncCommandAvailable {
    param([string]$CommandName)
    if (-not (Test-ADSyncModuleAvailable)) { return $false }
    $escaped = $CommandName.Replace("'", "''")
    $r = Invoke-ADSyncChild -ScriptText "if (Get-Command -Name '$escaped' -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"
    return [bool]$r.Success
}

function Test-CanReadLocalGroup {
    param([string]$GroupName)
    try {
        $null = net localgroup "$GroupName" 2>$null
        return $LASTEXITCODE -eq 0
    } catch {
        return $false
    }
}

function Ensure-Folder {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    return (Test-Path -LiteralPath $Path)
}

function Test-PathWritable {
    param([string]$Path)
    try {
        if (-not (Ensure-Folder -Path $Path)) { return $false }
        $testFile = Join-Path $Path ("write-test-{0}.tmp" -f ([guid]::NewGuid().Guid))
        'test' | Set-Content -LiteralPath $testFile -Encoding UTF8
        Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        return $false
    }
}

function New-ZipFromFolder {
    param(
        [string]$SourceFolder,
        [string]$ZipPath,
        [switch]$WhatIfOnly
    )

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        throw "Bronmap ontbreekt: $SourceFolder"
    }
    if ($WhatIfOnly) { return }
    if (Test-Path -LiteralPath $ZipPath) { Remove-Item -LiteralPath $ZipPath -Force }

    try {
        Compress-Archive -Path (Join-Path $SourceFolder '*') -DestinationPath $ZipPath -CompressionLevel Optimal -Force -ErrorAction Stop
    } catch {
        $rootMsg = $_.Exception.Message
        $fullMsg = ($_ | Out-String)
        if ($fullMsg -match 'being used by another process|Stream was not readable') {
            $stage = Join-Path ([System.IO.Path]::GetDirectoryName($ZipPath)) ("staging-{0}" -f ([guid]::NewGuid().Guid))
            New-Item -ItemType Directory -Path $stage -Force | Out-Null
            try {
                $null = & robocopy.exe $SourceFolder $stage /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NP
                $rc = $LASTEXITCODE
                if ($rc -ge 8) {
                    throw "Fallback staging-copy voor zip mislukte met robocopy-exitcode $rc. Oorspronkelijke fout: $rootMsg"
                }
                if (Test-Path -LiteralPath $ZipPath) {
                    Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
                }
                Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $ZipPath -CompressionLevel Optimal -Force -ErrorAction Stop
            } finally {
                Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
            }
        } else {
            throw
        }
    }
}

function Copy-LatestFiles {
    param(
        [string]$SourceFolder,
        [string[]]$Patterns,
        [string[]]$Destinations,
        [switch]$WhatIfOnly
    )

    $copied = @()
    foreach ($pattern in $Patterns) {
        $latest = Get-ChildItem -LiteralPath $SourceFolder -Filter $pattern -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if ($latest) {
            foreach ($dest in $Destinations) {
                if ([string]::IsNullOrWhiteSpace($dest)) { continue }
                if (-not $WhatIfOnly) {
                    Ensure-Folder -Path $dest | Out-Null
                    Copy-Item -LiteralPath $latest.FullName -Destination (Join-Path $dest $latest.Name) -Force
                }
            }
            $copied += $latest.FullName
        }
    }
    return $copied
}

function Save-ResultsReport {
    param([string]$Folder)
    try {
        if ([string]::IsNullOrWhiteSpace($Folder)) { return }
        Ensure-Folder -Path $Folder | Out-Null
        $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
        $csv  = Join-Path $Folder "EntraConnectUpgrade-Report-$ts.csv"
        $json = Join-Path $Folder "EntraConnectUpgrade-Report-$ts.json"
        $script:Results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $script:Results | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $json -Encoding UTF8
        Write-Log "Rapport opgeslagen: $csv" -Level 'Ok'
        Write-Log "Rapport opgeslagen: $json" -Level 'Ok'
    } catch {
        Write-Log "Rapport opslaan mislukt: $($_.Exception.Message)" -Level 'Error'
    }
}

# ---------------------------------------------------------------------------
# Validatie
# ---------------------------------------------------------------------------
function Validate-CoreEnvironment {
    $cfg          = Get-InputState
    $isAdmin      = Test-IsAdministrator
    $freeGB       = Get-FreeSpaceGB
    $hasModule    = Test-ADSyncModuleAvailable
    $service      = Get-ADSyncServiceSafe
    $canReadGroup = Test-CanReadLocalGroup -GroupName 'ADSyncAdmins'
    $syncPath     = Update-SyncPathLabel

    New-Result 'Stap 1' 'Documentatiemap voorbereiden' 'HANDMATIG' ("Maak in de klantdocumentatie (map Microsoft Entra Connect) een submap aan: Upgrade $(Get-Date -Format 'dd-MM-yyyy').")
    New-Result 'Stap 2'  'Monitoring in onderhoudsmodus' 'HANDMATIG' 'Blijft handmatig. Deze GUI bedient de monitoringtool niet.'
    New-Result 'Stap 2a' 'Beheeraccount tijdelijk Global Admin / Enterprise Admin' 'HANDMATIG' 'Blijft handmatig/JIT.'
    New-Result 'Stap 3'  'Aanmelden op de Entra Connect-server' 'HANDMATIG' 'Dit script valideert alleen de server waar je nu op draait.'

    if ($canReadGroup) {
        New-Result 'Stap 4' 'ADSyncAdmins-lidmaatschap controleren' 'OK' 'Lokale groep ADSyncAdmins is uitleesbaar.'
    } else {
        New-Result 'Stap 4' 'ADSyncAdmins-lidmaatschap controleren' 'LET OP' 'Lokale groep ADSyncAdmins kon niet worden gelezen. Controleer server en rechten.'
    }

    New-Result 'Stap 5' 'Laatste Entra Connect installer downloaden' 'HANDMATIG' 'Download blijft handmatig; vul daarna het pad in de GUI in.'

    if ($null -ne $freeGB -and $freeGB -ge 15) {
        New-Result 'Stap 6' 'Vrije ruimte C-schijf' 'OK' "Vrije ruimte op C: is $freeGB GB (minimaal 15 GB vereist)."
    } elseif ($null -ne $freeGB) {
        New-Result 'Stap 6' 'Vrije ruimte C-schijf' 'BLOKKADE' "Vrije ruimte op C: is slechts $freeGB GB. Minimaal 15 GB vereist."
    } else {
        New-Result 'Stap 6' 'Vrije ruimte C-schijf' 'LET OP' 'Vrije ruimte op C: kon niet worden vastgesteld.'
    }

    New-Result 'Stap 6b' 'Staging mode handmatig inschakelen' 'HANDMATIG' 'Blijft handmatig in de Entra Connect GUI.'

    if ($isAdmin) {
        New-Result 'Stap 7' 'PowerShell als Administrator' 'OK' 'De huidige sessie draait elevated.'
    } else {
        New-Result 'Stap 7' 'PowerShell als Administrator' 'BLOKKADE' 'Huidige sessie draait niet als Administrator.'
    }

    if ($hasModule) {
        New-Result 'Stap 8' 'ADSync module laden en scheduler bedienen' 'OK' 'ADSync module is aanwezig.'
    } else {
        New-Result 'Stap 8' 'ADSync module laden en scheduler bedienen' 'BLOKKADE' 'ADSync module niet gevonden.'
    }

    if ($service) {
        New-Result 'Stap 9' 'ADSync service controleren' 'OK' ('ADSync service bestaat en heeft status: ' + $service.Status)
    } else {
        New-Result 'Stap 9' 'ADSync service controleren' 'BLOKKADE' 'ADSync service niet gevonden. Waarschijnlijk niet de juiste server.'
    }

    $aadPath = 'C:\ProgramData\AADConnect'
    if (Test-Path -LiteralPath $aadPath) {
        New-Result 'Stap 10' 'AADConnect folder zippen en kopieren' 'OK' "Bronmap bestaat: $aadPath"
    } else {
        New-Result 'Stap 10' 'AADConnect folder zippen en kopieren' 'BLOKKADE' "Bronmap ontbreekt: $aadPath"
    }

    if ($syncPath) {
        New-Result 'Stap 11' 'Sync-installatiemap zippen (optioneel)' 'OK' "Sync-installatiemap gedetecteerd: $syncPath"
    } else {
        New-Result 'Stap 11' 'Sync-installatiemap zippen (optioneel)' 'INFO' 'Geen aparte sync-installatiemap gevonden. Dit is normaal bij nieuwere Entra Connect-versies; deze back-upstap wordt overgeslagen en blokkeert de upgrade niet.'
    }

    if ($service) {
        New-Result 'Stap 12' 'ADSync service starten' 'OK' 'Service bestaat; deze stap is technisch uitvoerbaar.'
    } else {
        New-Result 'Stap 12' 'ADSync service starten' 'BLOKKADE' 'Service ontbreekt.'
    }

    New-Result 'Stap 13-15' 'Configuratie exporteren (.json) en trace-log veiligstellen' 'HANDMATIG' 'Export Settings in de GUI blijft handmatig; daarna kopieert het script de nieuwste .json en trace-*.log.'

    if ([string]::IsNullOrWhiteSpace($cfg.LocalExportFolder)) {
        New-Result 'Stap 15' 'Lokale exportmap voor .json/.csv' 'LET OP' 'Geen lokale exportmap ingevuld.'
    } elseif (Test-PathWritable -Path $cfg.LocalExportFolder) {
        New-Result 'Stap 15' 'Lokale exportmap voor .json/.csv' 'OK' "Lokale exportmap is schrijfbaar: $($cfg.LocalExportFolder)"
    } else {
        New-Result 'Stap 15' 'Lokale exportmap voor .json/.csv' 'LET OP' "Lokale exportmap niet schrijfbaar/bereikbaar: $($cfg.LocalExportFolder)"
    }

    if (Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncExportDeletionThreshold') {
        $upnNote = if ([string]::IsNullOrWhiteSpace($cfg.AadUpn)) { 'Bij uitvoeren wordt om een Entra-account (UPN) gevraagd.' } else { "Entra-account voor deze stap: $($cfg.AadUpn)" }
        New-Result 'Stap 16' 'Deletion threshold controleren' 'OK' "Get-ADSyncExportDeletionThreshold beschikbaar. $upnNote"
    } else {
        New-Result 'Stap 16' 'Deletion threshold controleren' 'LET OP' 'Get-ADSyncExportDeletionThreshold niet beschikbaar.'
    }

    if (Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncConnector') {
        New-Result 'Stap 17' 'Connector export (.csv)' 'OK' 'Get-ADSyncConnector beschikbaar.'
    } else {
        New-Result 'Stap 17' 'Connector export (.csv)' 'LET OP' 'Get-ADSyncConnector niet beschikbaar.'
    }

    New-Result 'Stap 18-19' 'Licentiegroepen exporteren vanuit M365 Admin Center' 'HANDMATIG' 'Blijft handmatig. De GUI opent wel de map en het admin center.'
    New-Result 'Stap 20-25' 'Installer starten en upgrade-wizard doorlopen' 'HANDMATIG' 'Blijft handmatig: installer, Upgrade, inloggen, uitvoeren, Exit, staging mode uit.'

    if (Test-ADSyncCommandAvailable -CommandName 'Set-ADSyncScheduler') {
        New-Result 'Stap 26' 'Scheduler weer inschakelen' 'OK' 'Set-ADSyncScheduler beschikbaar.'
    } else {
        New-Result 'Stap 26' 'Scheduler weer inschakelen' 'LET OP' 'Set-ADSyncScheduler ontbreekt.'
    }

    if ((Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncAutoUpgrade') -or (Test-ADSyncCommandAvailable -CommandName 'Set-ADSyncAutoUpgrade')) {
        New-Result 'Stap 27' 'Auto Upgrade controleren/uitzetten' 'OK' 'AutoUpgrade-cmdlets aanwezig.'
    } else {
        New-Result 'Stap 27' 'Auto Upgrade controleren/uitzetten' 'LET OP' 'Niet alle AutoUpgrade-cmdlets zijn beschikbaar.'
    }

    New-Result 'Stap 28' 'Monitoring uit onderhoudsmodus' 'HANDMATIG' 'Blijft handmatig.'

    if (Test-ADSyncCommandAvailable -CommandName 'Start-ADSyncSyncCycle') {
        New-Result 'Stap 29' 'Delta sync starten' 'OK' 'Start-ADSyncSyncCycle beschikbaar.'
        New-Result 'Stap 30' 'Initial sync starten' 'OK' 'Runstatus- en sync-cmdlets grotendeels beschikbaar.'
    } else {
        New-Result 'Stap 29' 'Delta sync starten' 'LET OP' 'Cmdlets voor delta sync niet volledig beschikbaar.'
        New-Result 'Stap 30' 'Initial sync starten' 'LET OP' 'Cmdlets voor initial sync niet volledig beschikbaar.'
    }

    New-Result 'Stap 31' 'Entra Connect excelsheet bijwerken' 'HANDMATIG' 'Blijft handmatig.'
    New-Result 'Stap 32' 'Teammail versturen' 'HANDMATIG' 'Blijft handmatig.'

    if (Test-ADSyncCommandAvailable -CommandName 'Set-ADSyncScheduler') {
        New-Result 'Stap 33' 'Run history logging op 30 dagen zetten' 'OK' 'Set-ADSyncScheduler beschikbaar voor PurgeRunHistoryInterval.'
    } else {
        New-Result 'Stap 33' 'Run history logging op 30 dagen zetten' 'LET OP' 'Set-ADSyncScheduler ontbreekt.'
    }

    New-Result 'Stap 34' 'Tijdelijke rechten opruimen' 'HANDMATIG' 'Blijft handmatig/JIT-afhankelijk.'

    Write-Summary
}

# ---------------------------------------------------------------------------
# Handmatige checkpoints
# ---------------------------------------------------------------------------
function Invoke-ManualCheckpoint {
    param(
        [string]$Step,
        [string]$Title,
        [string]$Text
    )

    if (-not (Test-StepActive -StepId $Step)) {
        New-Result $Step $Title 'INFO' "Stap overgeslagen (herstart vanaf $script:StartFromStep)."
        return $true
    }

    $res = Show-StepDialog -Title "$Step - $Title" -Text $Text -Buttons OKCancel -Icon Information
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
        New-Result $Step $Title 'HANDMATIG' 'Gebruiker heeft handmatige actie bevestigd en doorgegaan.'
        return $true
    }
    New-Result $Step $Title 'LET OP' 'Gebruiker heeft de workflow onderbroken bij dit checkpoint.'
    return $false
}

# ---------------------------------------------------------------------------
# Dry run
# ---------------------------------------------------------------------------
function Invoke-WhatIfWorkflow {
    $cfg          = Get-InputState
    $isAdmin      = Test-IsAdministrator
    $freeGB       = Get-FreeSpaceGB
    $hasModule    = Test-ADSyncModuleAvailable
    $service      = Get-ADSyncServiceSafe
    $canReadGroup = Test-CanReadLocalGroup -GroupName 'ADSyncAdmins'
    $aadPath      = 'C:\ProgramData\AADConnect'
    $syncPath     = Update-SyncPathLabel

    Write-Log 'Dry-run werkt stap voor stap. Bij iedere handmatige stap verschijnt een bevestigingsvenster. Er worden geen wijzigingen doorgevoerd.'

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 1' -Title 'Documentatiemap voorbereiden' -Text @"
Ga nu naar de klantdocumentatie en open de map voor Microsoft Entra Connect.

Maak daar een nieuwe submap aan met de naam:
  Upgrade $(Get-Date -Format 'dd-MM-yyyy')

Klik op OK zodra de map is aangemaakt.
"@)) { return }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 2' -Title 'Monitoring in onderhoudsmodus' -Text 'Zet de server nu handmatig in onderhoudsmodus in de monitoringtool. Klik op OK zodra dit is gecontroleerd of gesimuleerd.')) { return }
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 2a' -Title 'Tijdelijke rechten toekennen' -Text 'Maak jezelf tijdelijk Global Administrator en Enterprise Administrator (handmatig/JIT). Klik op OK om verder te gaan.')) { return }
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 3' -Title 'RDP naar de juiste Entra Connect-server' -Text 'Bevestig dat je via de juiste route bent aangemeld op de echte Entra Connect server. Klik op OK om door te gaan.')) { return }

    if (Test-StepActive -StepId 'Stap 4') {
        if ($canReadGroup) {
            New-Result 'Stap 4' 'ADSyncAdmins-lidmaatschap controleren' 'OK' 'Lokale groep ADSyncAdmins is uitleesbaar.'
        } else {
            New-Result 'Stap 4' 'ADSyncAdmins-lidmaatschap controleren' 'LET OP' 'Lokale groep ADSyncAdmins kon niet worden gelezen.'
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 5' -Title 'Laatste Entra Connect installer downloaden' -Text @"
Zorg dat de juiste Entra Connect installer is gedownload en vul het pad in bij InstallerPath.

Huidig pad uit de GUI:
$($cfg.InstallerPath)

Klik op OK om verder te gaan.
"@)) { return }

    if (Test-StepActive -StepId 'Stap 5') {
        if ([string]::IsNullOrWhiteSpace($cfg.InstallerPath)) {
            New-Result 'Stap 5' 'InstallerPath valideren' 'LET OP' 'InstallerPath is leeg.'
        } elseif (Test-Path -LiteralPath $cfg.InstallerPath) {
            New-Result 'Stap 5' 'InstallerPath valideren' 'OK' "InstallerPath bestaat: $($cfg.InstallerPath)"
        } else {
            New-Result 'Stap 5' 'InstallerPath valideren' 'LET OP' "InstallerPath bestaat niet: $($cfg.InstallerPath)"
        }
    }

    if (Test-StepActive -StepId 'Stap 6') {
        if ($null -ne $freeGB -and $freeGB -ge 15) {
            New-Result 'Stap 6' 'Vrije ruimte C-schijf' 'OK' "Vrije ruimte op C: is $freeGB GB."
        } elseif ($null -ne $freeGB) {
            New-Result 'Stap 6' 'Vrije ruimte C-schijf' 'BLOKKADE' "Vrije ruimte op C: is slechts $freeGB GB. Minimaal 15 GB vereist."
        } else {
            New-Result 'Stap 6' 'Vrije ruimte C-schijf' 'LET OP' 'Vrije ruimte op C: kon niet worden vastgesteld.'
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 6b' -Title 'Staging mode handmatig inschakelen' -Text "Open de Entra Connect GUI en zet handmatig:`r`n- Staging mode AAN`r`n- synchronisatie bij config completion UIT`r`n`r`nKlik op OK om verder te gaan.")) { return }

    if (Test-StepActive -StepId 'Stap 7') {
        if ($isAdmin) {
            New-Result 'Stap 7' 'PowerShell als Administrator' 'OK' 'De huidige sessie draait elevated.'
        } else {
            New-Result 'Stap 7' 'PowerShell als Administrator' 'BLOKKADE' 'Huidige sessie draait niet als Administrator.'
        }
    }

    if (Test-StepActive -StepId 'Stap 8') {
        if ($hasModule) {
            New-Result 'Stap 8' 'ADSync module laden en scheduler bedienen' 'OK' 'ADSync module aanwezig. In een echte run zou de scheduler worden uitgezet.'
        } else {
            New-Result 'Stap 8' 'ADSync module laden en scheduler bedienen' 'BLOKKADE' 'ADSync module niet gevonden.'
        }
    }

    if (Test-StepActive -StepId 'Stap 9') {
        if ($service) {
            New-Result 'Stap 9' 'ADSync service controleren' 'OK' ('Status: ' + $service.Status + '. In een echte run wordt de service gestopt en later gestart.')
        } else {
            New-Result 'Stap 9' 'ADSync service controleren' 'BLOKKADE' 'ADSync service niet gevonden.'
        }
    }

    if (Test-StepActive -StepId 'Stap 10') {
        if (Test-Path -LiteralPath $aadPath) {
            New-Result 'Stap 10' 'AADConnect folder zippen en kopieren' 'OK' "Bronmap bestaat: $aadPath"
        } else {
            New-Result 'Stap 10' 'AADConnect folder zippen en kopieren' 'BLOKKADE' "Bronmap ontbreekt: $aadPath"
        }
    }

    if (Test-StepActive -StepId 'Stap 11') {
        if ($syncPath) {
            New-Result 'Stap 11' 'Sync-installatiemap zippen (optioneel)' 'OK' "In een echte run wordt een zip gemaakt van: $syncPath"
        } else {
            New-Result 'Stap 11' 'Sync-installatiemap zippen (optioneel)' 'INFO' 'Geen sync-installatiemap gevonden (normaal bij nieuwere versies). Deze back-upstap wordt overgeslagen.'
        }
    }

    if (Test-StepActive -StepId 'Stap 12') {
        if ($service) {
            New-Result 'Stap 12' 'ADSync service starten' 'OK' 'Service bestaat; kan na de back-up weer gestart worden.'
        } else {
            New-Result 'Stap 12' 'ADSync service starten' 'BLOKKADE' 'Service ontbreekt.'
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 13-15' -Title 'Configuratie exporteren (.json)' -Text @"
Voer nu handmatig in de Entra Connect GUI uit:

1. View or export current configuration
2. Exporteer de configuratie naar JSON
3. Rond dit venster af

Klik op OK om verder te gaan.
"@)) { return }

    if ([string]::IsNullOrWhiteSpace($cfg.LocalExportFolder)) {
        New-Result 'Stap 15' 'Lokale exportmap voor .json/.csv' 'LET OP' 'Geen lokale exportmap ingevuld.'
    } elseif (Test-PathWritable -Path $cfg.LocalExportFolder) {
        New-Result 'Stap 15' 'Lokale exportmap voor .json/.csv' 'OK' "Lokale exportmap is schrijfbaar: $($cfg.LocalExportFolder)"
    } else {
        New-Result 'Stap 15' 'Lokale exportmap voor .json/.csv' 'LET OP' "Lokale exportmap niet schrijfbaar: $($cfg.LocalExportFolder)"
    }

    if (Test-StepActive -StepId 'Stap 16') {
        if (Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncExportDeletionThreshold') {
            New-Result 'Stap 16' 'Deletion threshold controleren' 'OK' 'Cmdlet beschikbaar. In een echte run vraagt de GUI om een Entra-account (UPN) en verschijnt mogelijk een aanmeldvenster.'
        } else {
            New-Result 'Stap 16' 'Deletion threshold controleren' 'LET OP' 'Cmdlet niet beschikbaar.'
        }
    }

    if (Test-StepActive -StepId 'Stap 17') {
        if (Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncConnector') {
            New-Result 'Stap 17' 'Connector export (.csv)' 'OK' 'Get-ADSyncConnector beschikbaar.'
        } else {
            New-Result 'Stap 17' 'Connector export (.csv)' 'LET OP' 'Get-ADSyncConnector niet beschikbaar.'
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 18-19' -Title 'Licentiegroep-exports (M365)' -Text @"
Voer nu handmatig de licentiegroep-export(s) uit vanuit Microsoft 365 Admin Center.

LocalExportFolder uit de GUI:
$($cfg.LocalExportFolder)

Klik op OK om verder te gaan.
"@)) { return }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 20-25' -Title 'Installer starten en wizard doorlopen' -Text @"
Nu zou je handmatig de Entra Connect update setup starten.

Pad uit de GUI:
$($cfg.InstallerPath)

Wizard: Upgrade kiezen > Inloggen > Upgrade uitvoeren > Exit > Staging mode UIT.

Klik op OK om verder te gaan.
"@)) { return }

    if (Test-StepActive -StepId 'Stap 26') {
        if (Test-ADSyncCommandAvailable -CommandName 'Set-ADSyncScheduler') {
            New-Result 'Stap 26' 'Scheduler weer inschakelen' 'OK' 'Set-ADSyncScheduler beschikbaar.'
        } else {
            New-Result 'Stap 26' 'Scheduler weer inschakelen' 'LET OP' 'Set-ADSyncScheduler ontbreekt.'
        }
    }

    if (Test-StepActive -StepId 'Stap 27') {
        if ((Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncAutoUpgrade') -or (Test-ADSyncCommandAvailable -CommandName 'Set-ADSyncAutoUpgrade')) {
            New-Result 'Stap 27' 'Auto Upgrade controleren/uitzetten' 'OK' 'AutoUpgrade-cmdlets aanwezig.'
        } else {
            New-Result 'Stap 27' 'Auto Upgrade controleren/uitzetten' 'LET OP' 'AutoUpgrade-cmdlets niet volledig beschikbaar.'
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 28' -Title 'Monitoring uit onderhoudsmodus' -Text 'Haal de server handmatig uit onderhoudsmodus in de monitoringtool. Klik op OK om verder te gaan.')) { return }

    if (Test-StepActive -StepId 'Stap 29') {
        if (Test-ADSyncCommandAvailable -CommandName 'Start-ADSyncSyncCycle') {
            New-Result 'Stap 29' 'Delta sync starten' 'OK' 'Start-ADSyncSyncCycle beschikbaar.'
        } else {
            New-Result 'Stap 29' 'Delta sync starten' 'LET OP' 'Cmdlets niet volledig beschikbaar.'
        }
    }

    if (Test-StepActive -StepId 'Stap 30') {
        if (Test-ADSyncCommandAvailable -CommandName 'Start-ADSyncSyncCycle') {
            New-Result 'Stap 30' 'Initial sync starten' 'OK' 'Runstatus- en sync-cmdlets beschikbaar.'
        } else {
            New-Result 'Stap 30' 'Initial sync starten' 'LET OP' 'Cmdlets niet volledig beschikbaar.'
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 31' -Title 'Excelsheet bijwerken' -Text 'Werk na afloop handmatig de centrale Entra Connect excelsheet bij. Klik op OK om verder te gaan.')) { return }
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 32' -Title 'Teammail versturen' -Text 'Verstuur na afloop handmatig de teammail. Klik op OK om verder te gaan.')) { return }

    if (Test-StepActive -StepId 'Stap 33') {
        if (Test-ADSyncCommandAvailable -CommandName 'Set-ADSyncScheduler') {
            New-Result 'Stap 33' 'Run history op 30 dagen' 'OK' 'Set-ADSyncScheduler beschikbaar voor PurgeRunHistoryInterval.'
        } else {
            New-Result 'Stap 33' 'Run history op 30 dagen' 'LET OP' 'Set-ADSyncScheduler ontbreekt.'
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 34' -Title 'Tijdelijke rechten opruimen' -Text 'Verwijder na afloop de tijdelijke Global Admin / Enterprise Admin rechten. Klik op OK om de dry-run af te ronden.')) { return }

    Write-Summary
}

# ---------------------------------------------------------------------------
# Uitvoeren
# ---------------------------------------------------------------------------
function Invoke-ExecuteWorkflow {
    $cfg = Get-InputState
    $syncPath = Update-SyncPathLabel

    if (-not [string]::IsNullOrWhiteSpace($script:StartFromStep)) {
        Write-Log "Herstart workflow vanaf: $script:StartFromStep - eerdere stappen worden overgeslagen." -Level 'Warn'
    }

    if (Test-StepActive -StepId 'Stap 7') {
        if (-not (Test-IsAdministrator)) {
            New-Result 'Stap 7' 'PowerShell als Administrator' 'BLOKKADE' 'Start dit script opnieuw als Administrator.'
            return
        }
    }

    if (Test-StepActive -StepId 'Stap 8') {
        if (-not (Test-ADSyncModuleAvailable)) {
            New-Result 'Stap 8' 'ADSync module laden' 'BLOKKADE' 'ADSync module niet gevonden.'
            return
        }
    }

    if (Test-StepActive -StepId 'Stap 9') {
        $service = Get-ADSyncServiceSafe
        if (-not $service) {
            New-Result 'Stap 9' 'ADSync service controleren' 'BLOKKADE' 'ADSync service niet gevonden. Dit lijkt niet de juiste server.'
            return
        }
    }

    # Alleen de configuratiemap is nog een harde eis. De sync-installatiemap is
    # optioneel geworden omdat nieuwere Entra Connect-versies deze niet meer
    # op het klassieke pad plaatsen.
    if (Test-StepActive -StepId 'Stap 10') {
        if (-not (Test-Path -LiteralPath 'C:\ProgramData\AADConnect')) {
            New-Result 'Precheck' 'Benodigde bronmap ontbreekt' 'BLOKKADE' 'C:\ProgramData\AADConnect'
            return
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($cfg.LocalExportFolder) -and -not (Test-PathWritable -Path $cfg.LocalExportFolder)) {
        New-Result 'Stap 15' 'LocalExportFolder' 'LET OP' "LocalExportFolder lijkt niet schrijfbaar: $($cfg.LocalExportFolder)"
    }

    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $shortName = if ([string]::IsNullOrWhiteSpace($cfg.CustomerName)) { 'Customer' } else { ($cfg.CustomerName -replace '[^a-zA-Z0-9_-]', '_') }
    $workFolder = Join-Path $env:TEMP ("EntraConnectUpgrade-$shortName-$ts")
    Ensure-Folder -Path $workFolder | Out-Null

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 1' -Title 'Documentatiemap voorbereiden' -Text "Ga nu naar de klantdocumentatie en open de map voor Microsoft Entra Connect.`r`n`r`nMaak daar een nieuwe submap aan met de naam:`r`n  Upgrade $(Get-Date -Format 'dd-MM-yyyy')`r`n`r`nKlik op OK zodra de map is aangemaakt.")) { return }
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 2' -Title 'Monitoring in onderhoudsmodus' -Text 'Zet de server nu handmatig in onderhoudsmodus in de monitoringtool. Klik op OK als dit gereed is.')) { return }
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 2a' -Title 'Tijdelijke rechten toekennen' -Text 'Maak jezelf nu tijdelijk Global Administrator en Enterprise Administrator. Klik op OK als dit gereed is.')) { return }
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 3' -Title 'Bevestig juiste server' -Text 'Bevestig dat je nu op de juiste Entra Connect server bent aangemeld. Klik op OK om verder te gaan.')) { return }
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 5' -Title 'Laatste installer' -Text 'Zorg dat de juiste Entra Connect installer is gedownload en vul het pad bij InstallerPath in. Klik daarna op OK.')) { return }

    if (Test-StepActive -StepId 'Stap 5') {
        if ([string]::IsNullOrWhiteSpace($cfg.InstallerPath) -or -not (Test-Path -LiteralPath $cfg.InstallerPath)) {
            New-Result 'Stap 5' 'InstallerPath valideren' 'BLOKKADE' 'InstallerPath is leeg of bestaat niet.'
            return
        }
    }

    if (Test-StepActive -StepId 'Stap 6') {
        $freeGB = Get-FreeSpaceGB
        if ($null -eq $freeGB -or $freeGB -lt 15) {
            $detected = if ($null -eq $freeGB) { 'onbekend' } else { [string]$freeGB }
            New-Result 'Stap 6' 'Vrije ruimte C-schijf' 'BLOKKADE' "Onvoldoende vrije ruimte op C:. Gedetecteerd: $detected GB"
            return
        }
        New-Result 'Stap 6' 'Vrije ruimte C-schijf' 'OK' "Vrije ruimte op C: $freeGB GB"
    }

    if (Test-StepActive -StepId 'Stap 8') {
        # Tijdens een lopende sync-cyclus weigert de engine configuratiewijzigingen.
        if (-not (Wait-ADSyncIdle -StepLabel 'Stap 8')) {
            New-Result 'Stap 8' 'Scheduler uitzetten' 'BLOKKADE' 'Er liep nog een synchronisatie en wachten is afgebroken. Start de workflow opnieuw zodra de sync klaar is.'
            return
        }
        $rSchedulerOff = Invoke-ADSyncChild -ScriptText 'Set-ADSyncScheduler -SyncCycleEnabled $false' -WaitMessage 'Stap 8 - scheduler uitzetten'
        if (-not $rSchedulerOff.Success) {
            New-Result 'Stap 8' 'Scheduler uitzetten' 'BLOKKADE' (Get-ADSyncErrorSummary $rSchedulerOff.Output)
            return
        }
        New-Result 'Stap 8' 'Scheduler uitzetten' 'OK' 'Set-ADSyncScheduler -SyncCycleEnabled $false uitgevoerd.'
    }

    if (Test-StepActive -StepId 'Stap 9') {
        $script:ServiceStoppedByScript = $true
        Stop-Service -Name ADSync -Force
        New-Result 'Stap 9' 'ADSync service stoppen' 'OK' 'Service ADSync is gestopt.'
    }

    if (Test-StepActive -StepId 'Stap 10') {
        $zipAad = Join-Path $workFolder ("AADConnect-$shortName-$ts.zip")
        New-ZipFromFolder -SourceFolder 'C:\ProgramData\AADConnect' -ZipPath $zipAad
        if (-not [string]::IsNullOrWhiteSpace($cfg.LocalExportFolder)) {
            Ensure-Folder -Path $cfg.LocalExportFolder | Out-Null
            Copy-Item -LiteralPath $zipAad -Destination (Join-Path $cfg.LocalExportFolder (Split-Path $zipAad -Leaf)) -Force
        }
        New-Result 'Stap 10' 'AADConnect folder zippen en kopieren' 'OK' "Backup gemaakt en gekopieerd: $zipAad"
    }

    if (Test-StepActive -StepId 'Stap 11') {
        if ($syncPath -and (Test-Path -LiteralPath $syncPath)) {
            $zipSync = $null
            try {
                $zipSync = Join-Path $workFolder ("SyncEngine-$shortName-$ts.zip")
                New-ZipFromFolder -SourceFolder $syncPath -ZipPath $zipSync
                if (-not [string]::IsNullOrWhiteSpace($cfg.LocalExportFolder)) {
                    Ensure-Folder -Path $cfg.LocalExportFolder | Out-Null
                    Copy-Item -LiteralPath $zipSync -Destination (Join-Path $cfg.LocalExportFolder (Split-Path $zipSync -Leaf)) -Force
                }
                New-Result 'Stap 11' 'Sync-installatiemap zippen (optioneel)' 'OK' "Backup gemaakt van $syncPath -> $zipSync"
            } catch {
                if ($zipSync -and (Test-Path -LiteralPath $zipSync)) {
                    Remove-Item -LiteralPath $zipSync -Force -ErrorAction SilentlyContinue
                }
                New-Result 'Stap 11' 'Sync-installatiemap zippen (optioneel)' 'LET OP' ("Back-up van de sync-installatiemap is niet gelukt; onvolledige zip is verwijderd. De upgrade kan doorgaan, want deze map wordt door de installer opnieuw geplaatst. " +
                    "Een bestand werd vastgehouden door een ander proces (bijv. een open Synchronization Service Manager, een Entra Connect Health-agent of een PowerShell-sessie met de ADSync-module). Fout: $($_.Exception.Message)")
            }
        } else {
            New-Result 'Stap 11' 'Sync-installatiemap zippen (optioneel)' 'INFO' 'Geen sync-installatiemap gevonden. Bij nieuwere Entra Connect-versies bestaat deze map niet meer; stap overgeslagen.'
        }
    }

    if (Test-StepActive -StepId 'Stap 12') {
        Start-Service -Name ADSync
        $script:ServiceStoppedByScript = $false
        New-Result 'Stap 12' 'ADSync service starten' 'OK' 'Service ADSync is gestart.'
    }

    $manualConfigText = @"
Voer nu handmatig in de Entra Connect GUI uit:

1. View or export current configuration
2. Exporteer de configuratie naar JSON
3. Rond dit venster af

Sla de JSON op in C:\ProgramData\AADConnect of een andere bekende locatie.
Klik daarna op OK. Het script kopieert vervolgens de nieuwste JSON en trace-log.
"@
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 13-15' -Title 'Configuratie exporteren (.json)' -Text $manualConfigText)) { return }

    $destinations = @()
    if (-not [string]::IsNullOrWhiteSpace($cfg.LocalExportFolder)) { $destinations += $cfg.LocalExportFolder }
    $copied = @(Copy-LatestFiles -SourceFolder 'C:\ProgramData\AADConnect' -Patterns @('*.json', 'trace-*.log') -Destinations $destinations)
    if ($copied.Count -gt 0) {
        New-Result 'Stap 15' 'JSON en trace-log veiligstellen' 'OK' ('Gekopieerd: ' + ($copied -join '; '))
    } else {
        New-Result 'Stap 15' 'JSON en trace-log veiligstellen' 'LET OP' 'Geen recente .json of trace-*.log gevonden in C:\ProgramData\AADConnect.'
    }

    if (Test-StepActive -StepId 'Stap 16') {
        if (Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncExportDeletionThreshold') {
            Set-StatusText 'Stap 16 - wacht op Entra-account (UPN)...'
            $upn = Show-AadUpnPrompt -Default $cfg.AadUpn
            if ([string]::IsNullOrWhiteSpace($upn)) {
                New-Result 'Stap 16' 'Deletion threshold controleren' 'LET OP' 'Overgeslagen door gebruiker. Controleer handmatig: Get-ADSyncExportDeletionThreshold -AADUserName "<UPN>".'
            } else {
                # Onthouden voor een eventuele herstart van de workflow.
                $script:TxtAadUpn.Text = $upn
                try {
                    $safeUpn = $upn.Replace("'", "''")
                    $thresholdResult = Invoke-ADSyncChild `
                        -ScriptText "Get-ADSyncExportDeletionThreshold -AADUserName '$safeUpn' | Format-List * | Out-String -Width 4096" `
                        -TimeoutSeconds 600 `
                        -WaitMessage "Stap 16 - wacht op aanmelding als $upn"
                    if ($thresholdResult.TimedOut) { throw "Geen reactie binnen 10 minuten (aanmeldvenster niet afgerond?). $($thresholdResult.Output)" }
                    if (-not $thresholdResult.Success) { throw (Get-ADSyncErrorSummary $thresholdResult.Output) }
                    $threshold = $thresholdResult.Output
                    $thresholdFileBase = "DeletionThreshold-$shortName-$ts"
                    foreach ($dest in $destinations) {
                        if ([string]::IsNullOrWhiteSpace($dest)) { continue }
                        Ensure-Folder -Path $dest | Out-Null
                        $out = Join-Path $dest "$thresholdFileBase.txt"
                        $threshold | Set-Content -LiteralPath $out -Encoding UTF8
                    }
                    $compact = ($threshold -replace '\s+', ' ').Trim()
                    if ($compact.Length -gt 300) { $compact = $compact.Substring(0, 300) + '...' }
                    New-Result 'Stap 16' 'Deletion threshold controleren' 'OK' "Uitgelezen als $upn en vastgelegd. $compact"
                } catch {
                    New-Result 'Stap 16' 'Deletion threshold controleren' 'LET OP' $_.Exception.Message
                }
            }
        } else {
            New-Result 'Stap 16' 'Deletion threshold controleren' 'LET OP' 'Get-ADSyncExportDeletionThreshold ontbreekt.'
        }
    }

    if (Test-StepActive -StepId 'Stap 17') {
        try {
            if (-not (Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncConnector')) { throw 'Get-ADSyncConnector cmdlet ontbreekt.' }
            $csvs = @()
            foreach ($dest in $destinations) {
                if ([string]::IsNullOrWhiteSpace($dest)) { continue }
                Ensure-Folder -Path $dest | Out-Null
                $outFile = Join-Path $dest ("ADSyncConnector-$shortName-$ts.csv")
                $safeOut = $outFile.Replace("'", "''")
                $exportResult = Invoke-ADSyncChild -ScriptText "Get-ADSyncConnector | Select-Object * | Export-Csv -LiteralPath '$safeOut' -NoTypeInformation -Encoding UTF8 -Delimiter ';'"
                if (-not $exportResult.Success) { throw (Get-ADSyncErrorSummary $exportResult.Output) }
                $csvs += $outFile
            }
            New-Result 'Stap 17' 'Connector export (.csv)' 'OK' ('Connector-export weggeschreven: ' + ($csvs -join '; '))
        } catch {
            New-Result 'Stap 17' 'Connector export (.csv)' 'LET OP' $_.Exception.Message
        }
    }

    if ($cfg.OpenAdminCenter) {
        try { Start-Process 'https://admin.microsoft.com/Adminportal/Home?#/licenses' } catch { }
    }
    if (-not [string]::IsNullOrWhiteSpace($cfg.LocalExportFolder) -and (Test-Path -LiteralPath $cfg.LocalExportFolder)) {
        try { Start-Process explorer.exe $cfg.LocalExportFolder } catch { }
    }
    if (-not (Invoke-ManualCheckpoint -Step 'Stap 18-19' -Title 'Licentiegroep-exports handmatig uitvoeren' -Text 'Voer nu handmatig de licentiegroep-export(s) uit vanuit Microsoft 365 Admin Center en sla de CSV-bestanden op in de LocalExportFolder. Klik op OK zodra dit gereed is.')) { return }

    if (Test-StepActive -StepId 'Stap 20-25') {
        $installText = @"
Nu wordt de Entra Connect update setup gestart.

Pad dat in de GUI is ingevuld:
$($cfg.InstallerPath)

Doorloop daarna zelf de wizard:
- Upgrade kiezen
- Inloggen
- Upgrade uitvoeren
- Exit
- Daarna staging mode weer UIT zetten

Klik op OK om de installer te starten.
"@
        $r = Show-StepDialog -Title 'Stap 20-25 - Handmatige Entra Connect setup' -Text $installText -Buttons OKCancel -Icon Warning
        if ($r -ne [System.Windows.Forms.DialogResult]::OK) {
            New-Result 'Stap 20-25' 'Entra Connect setup' 'LET OP' 'Gebruiker heeft de setup-start geannuleerd.'
            return
        }

        Start-Process -FilePath $cfg.InstallerPath
        New-Result 'Stap 20-25' 'Entra Connect setup' 'HANDMATIG' 'Installer is gestart. Gebruiker doorloopt nu zelf de wizard.'

        if (-not (Invoke-ManualCheckpoint -Step 'Stap 20-25' -Title 'Bevestig afgeronde wizard' -Text "De installer is gestart. Voer nu de volgende stappen handmatig uit:`r`n`r`n1. Doorloop de upgrade-wizard (Upgrade > Inloggen > Uitvoeren > Exit)`r`n2. Start daarna de Entra Connect wizard OPNIEUW op`r`n3. Ga naar Configure en schakel Staging mode handmatig UIT`r`n4. Sluit de wizard af`r`n`r`nKlik pas op OK als alle stappen volledig zijn afgerond.")) { return }
    }

    if (Test-StepActive -StepId 'Stap 26') {
        # Run history (stap 33) alvast instellen zolang de scheduler nog UIT staat.
        # Zodra de scheduler aan gaat, start die vrijwel direct een sync-cyclus en
        # weigert Set-ADSyncScheduler wijzigingen ('Configuration can not be changed').
        $script:PurgeSetEarly = $false
        $earlyPurge = Set-ADSyncRunHistory30 -StepLabel 'Stap 26'
        if ($earlyPurge.Success) {
            $script:PurgeSetEarly = $true
            Write-Log 'Stap 26: run history alvast op 30 dagen gezet (onderdeel van stap 33, vooruitgetrokken omdat de scheduler nog uit stond).' -Level 'Ok'
        } else {
            Write-Log "Stap 26: run history nog niet ingesteld, wordt bij stap 33 opnieuw geprobeerd. $($earlyPurge.Detail)" -Level 'Warn'
        }

        $schedulerOk = $false
        $schedulerMsg = ''
        for ($attempt = 1; $attempt -le 3 -and -not $schedulerOk; $attempt++) {
            if (-not (Wait-ADSyncIdle -StepLabel 'Stap 26')) { $schedulerMsg = 'Overgeslagen: er liep nog een synchronisatie.'; break }
            $rSchedulerOn = Invoke-ADSyncChild -ScriptText 'Set-ADSyncScheduler -SyncCycleEnabled $true' -WaitMessage 'Stap 26 - scheduler inschakelen'
            if ($rSchedulerOn.Success) { $schedulerOk = $true; break }
            $schedulerMsg = Get-ADSyncErrorSummary $rSchedulerOn.Output
            if ($rSchedulerOn.Output -notmatch 'Configuration can not be changed|Cannot change configuration') { break }
            Wait-Responsive -Seconds 10 -Message 'Stap 26 - korte pauze voor nieuwe poging'
        }
        if ($schedulerOk) {
            New-Result 'Stap 26' 'Scheduler weer inschakelen' 'OK' 'Set-ADSyncScheduler -SyncCycleEnabled $true uitgevoerd. De scheduler start nu zelf binnen enkele minuten een sync-cyclus.'
        } else {
            New-Result 'Stap 26' 'Scheduler weer inschakelen' 'LET OP' "$schedulerMsg Voer handmatig uit: Set-ADSyncScheduler -SyncCycleEnabled `$true"
        }
    }

    if (Test-StepActive -StepId 'Stap 27') {
        if ($cfg.DisableAutoUpgrade -and (Test-ADSyncCommandAvailable -CommandName 'Get-ADSyncAutoUpgrade') -and (Test-ADSyncCommandAvailable -CommandName 'Set-ADSyncAutoUpgrade')) {
            try {
                $autoResult = Invoke-ADSyncChild -ScriptText 'Get-ADSyncAutoUpgrade | Out-String'
                if (-not $autoResult.Success) { throw (Get-ADSyncErrorSummary $autoResult.Output) }
                if ($autoResult.Output -notmatch 'Disabled') {
                    $setAuto = Invoke-ADSyncChild -ScriptText 'Set-ADSyncAutoUpgrade -AutoUpgradeState Disabled'
                    if (-not $setAuto.Success) { throw (Get-ADSyncErrorSummary $setAuto.Output) }
                    New-Result 'Stap 27' 'Auto Upgrade uitzetten' 'OK' 'AutoUpgrade is op Disabled gezet.'
                } else {
                    New-Result 'Stap 27' 'Auto Upgrade uitzetten' 'OK' 'AutoUpgrade stond al op Disabled.'
                }
            } catch {
                New-Result 'Stap 27' 'Auto Upgrade uitzetten' 'LET OP' $_.Exception.Message
            }
        } else {
            New-Result 'Stap 27' 'Auto Upgrade uitzetten' 'LET OP' 'AutoUpgrade-cmdlets ontbreken of de optie staat uit in de GUI.'
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 28' -Title 'Monitoring uit onderhoudsmodus' -Text 'Haal de server nu handmatig uit onderhoudsmodus in de monitoringtool. Klik op OK als dit gereed is.')) { return }

    if (Test-StepActive -StepId 'Stap 29') {
        if ($cfg.StartDelta -and (Test-ADSyncCommandAvailable -CommandName 'Start-ADSyncSyncCycle')) {
            $r29 = Start-ADSyncCycleWhenIdle -PolicyType 'Delta' -StepLabel 'Stap 29'
            New-Result 'Stap 29' 'Delta sync starten' $r29.Status $r29.Detail
        } else {
            if (-not (Invoke-ManualCheckpoint -Step 'Stap 29' -Title 'Delta sync handmatig uitvoeren' -Text 'Delta sync is niet automatisch gestart (optie staat uit in de GUI). Voer deze stap nu handmatig uit indien vereist. Klik op OK om verder te gaan.')) { return }
        }
    }

    if (Test-StepActive -StepId 'Stap 30') {
        if ($cfg.StartInitial -and (Test-ADSyncCommandAvailable -CommandName 'Start-ADSyncSyncCycle')) {
            $r30 = Start-ADSyncCycleWhenIdle -PolicyType 'Initial' -StepLabel 'Stap 30'
            New-Result 'Stap 30' 'Initial sync starten' $r30.Status $r30.Detail
        } else {
            New-Result 'Stap 30' 'Initial sync starten' 'HANDMATIG' 'Initial sync niet automatisch gestart volgens GUI-opties.'
        }
    }

    if (Test-StepActive -StepId 'Stap 33') {
        try {
            $svc33 = Get-ADSyncServiceSafe
            if ($null -eq $svc33) { throw 'ADSync service niet gevonden. Kan PurgeRunHistoryInterval niet instellen.' }
            if ($svc33.Status -ne 'Running') {
                Write-Log 'Stap 33: ADSync service is niet actief - service wordt gestart.' -Level 'Warn'
                Start-Service -Name ADSync -ErrorAction Stop
                Wait-Responsive -Seconds 5 -Message 'Stap 33 - ADSync service starten'
            }

            # Eerst controleren: bij stap 26 is dit normaal al ingesteld.
            $currentDays = Get-ADSyncPurgeIntervalDays
            if ($null -ne $currentDays -and [math]::Abs($currentDays - 30) -lt 0.01) {
                $how = if ($script:PurgeSetEarly) { 'ingesteld bij stap 26, voordat de scheduler aan ging' } else { 'stond al goed' }
                New-Result 'Stap 33' 'Run history op 30 dagen' 'OK' "PurgeRunHistoryInterval staat op 30 dagen ($how)."
            } else {
                $purge = Set-ADSyncRunHistory30 -StepLabel 'Stap 33'
                $st = if ($purge.Success) { 'OK' } else { 'LET OP' }
                New-Result 'Stap 33' 'Run history op 30 dagen' $st $purge.Detail
            }

            # Bij een herstart vanaf stap 33 kan de scheduler nog uit staan.
            $schedState = Invoke-ADSyncChild -ScriptText 'Write-Output ("SYNCENABLED=" + (Get-ADSyncScheduler).SyncCycleEnabled)' -TimeoutSeconds 120
            if ($schedState.Output -match '(?m)^SYNCENABLED=False') {
                Write-Log 'Stap 33: scheduler stond uit - wordt ingeschakeld.' -Level 'Warn'
                $enableResult = Invoke-ADSyncChild -ScriptText 'Set-ADSyncScheduler -SyncCycleEnabled $true'
                if (-not $enableResult.Success) {
                    Write-Log "Stap 33: scheduler inschakelen mislukt: $(Get-ADSyncErrorSummary $enableResult.Output)" -Level 'Warn'
                }
            }
        } catch {
            New-Result 'Stap 33' 'Run history op 30 dagen' 'LET OP' $_.Exception.Message
        }
    }

    if (-not (Invoke-ManualCheckpoint -Step 'Stap 34' -Title 'Tijdelijke rechten opruimen' -Text 'Verwijder nu de tijdelijke Global Admin / Enterprise Admin rechten. Klik op OK zodra dit gereed is.')) { return }

    Write-Summary
    Save-ResultsReport -Folder $cfg.LocalExportFolder
}

# ===========================================================================
#  UI
# ===========================================================================
function Get-RoundedPath {
    param([System.Drawing.Rectangle]$Rect, [int]$Radius = 8)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $Radius * 2
    $path.AddArc($Rect.X, $Rect.Y, $d, $d, 180, 90)
    $path.AddArc($Rect.Right - $d - 1, $Rect.Y, $d, $d, 270, 90)
    $path.AddArc($Rect.Right - $d - 1, $Rect.Bottom - $d - 1, $d, $d, 0, 90)
    $path.AddArc($Rect.X, $Rect.Bottom - $d - 1, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-Card {
    param([int]$X, [int]$Y, [int]$W, [int]$H, [string]$Anchor = 'Top, Left')
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location  = New-Object System.Drawing.Point($X, $Y)
    $panel.Size      = New-Object System.Drawing.Size($W, $H)
    $panel.BackColor = $script:Clr.Card
    $panel.Anchor    = [System.Windows.Forms.AnchorStyles]$Anchor
    $panel.Add_Paint({
        param($sender, $e)
        $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $rect = New-Object System.Drawing.Rectangle(0, 0, ($sender.Width - 1), ($sender.Height - 1))
        $path = Get-RoundedPath -Rect $rect -Radius 8
        $pen  = New-Object System.Drawing.Pen($script:Clr.Border, 1)
        $e.Graphics.DrawPath($pen, $path)
        $pen.Dispose(); $path.Dispose()
    })
    return $panel
}

function New-SectionLabel {
    param($Parent, [string]$Text, [int]$X, [int]$Y)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Text.ToUpperInvariant()
    $lbl.Font      = $script:Fonts.Caption
    $lbl.ForeColor = $script:Clr.Muted
    $lbl.Location  = New-Object System.Drawing.Point($X, $Y)
    $lbl.AutoSize  = $true
    $Parent.Controls.Add($lbl)
    return $lbl
}

function New-ModernButton {
    param(
        [string]$Text,
        [int]$X, [int]$Y, [int]$W, [int]$H,
        [ValidateSet('Primary', 'Secondary', 'Danger')]
        [string]$Kind = 'Secondary',
        [string]$Anchor = 'Top, Left'
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = $Text
    $btn.Location  = New-Object System.Drawing.Point($X, $Y)
    $btn.Size      = New-Object System.Drawing.Size($W, $H)
    $btn.FlatStyle = 'Flat'
    $btn.Font      = $script:Fonts.Bold
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btn.Anchor    = [System.Windows.Forms.AnchorStyles]$Anchor
    $btn.FlatAppearance.BorderSize = 1

    switch ($Kind) {
        'Primary' {
            $btn.BackColor = $script:Clr.Accent
            $btn.ForeColor = [System.Drawing.Color]::White
            $btn.FlatAppearance.BorderColor        = $script:Clr.Accent
            $btn.FlatAppearance.MouseOverBackColor = $script:Clr.AccentDark
            $btn.FlatAppearance.MouseDownBackColor = $script:Clr.AccentDark
        }
        'Danger' {
            $btn.BackColor = $script:Clr.Danger
            $btn.ForeColor = [System.Drawing.Color]::White
            $btn.FlatAppearance.BorderColor        = $script:Clr.Danger
            $btn.FlatAppearance.MouseOverBackColor = $script:Clr.DangerDark
            $btn.FlatAppearance.MouseDownBackColor = $script:Clr.DangerDark
        }
        default {
            $btn.BackColor = [System.Drawing.Color]::White
            $btn.ForeColor = $script:Clr.Text
            $btn.FlatAppearance.BorderColor        = $script:Clr.Border
            $btn.FlatAppearance.MouseOverBackColor = $script:Clr.AccentLight
            $btn.FlatAppearance.MouseDownBackColor = $script:Clr.AccentLight
        }
    }
    return $btn
}

function Add-LabeledTextBox {
    param($Parent, [string]$Label, [int]$X, [int]$Y, [int]$Width, [string]$Anchor = 'Top, Left')
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Label
    $lbl.Font      = $script:Fonts.Caption
    $lbl.ForeColor = $script:Clr.Muted
    $lbl.Location  = New-Object System.Drawing.Point($X, $Y)
    $lbl.AutoSize  = $true
    $Parent.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location    = New-Object System.Drawing.Point($X, ($Y + 17))
    $txt.Width       = $Width
    $txt.BorderStyle = 'FixedSingle'
    $txt.Font        = $script:Fonts.Base
    $txt.Anchor      = [System.Windows.Forms.AnchorStyles]$Anchor
    $Parent.Controls.Add($txt)
    return $txt
}

function New-ModernCheckBox {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [bool]$Checked = $false, [string]$Anchor = 'Top, Left')
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text      = $Text
    $chk.Location  = New-Object System.Drawing.Point($X, $Y)
    $chk.AutoSize  = $true
    $chk.Checked   = $Checked
    $chk.Font      = $script:Fonts.Base
    $chk.ForeColor = $script:Clr.Text
    $chk.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $chk.Anchor    = [System.Windows.Forms.AnchorStyles]$Anchor
    $Parent.Controls.Add($chk)
    return $chk
}

# --- Form ------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Entra Connect Upgrade - Controller'
$form.StartPosition = 'CenterScreen'
$form.Size          = New-Object System.Drawing.Size(1420, 900)
$form.MinimumSize   = New-Object System.Drawing.Size(1180, 760)
$form.BackColor     = $script:Clr.Bg
$form.Font          = $script:Fonts.Base
$form.ForeColor     = $script:Clr.Text

# --- Header ----------------------------------------------------------------
$header = New-Object System.Windows.Forms.Panel
$header.Dock      = 'Top'
$header.Height    = 66
$header.BackColor = $script:Clr.Header
$form.Controls.Add($header)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = 'Entra Connect Upgrade'
$lblTitle.Font      = $script:Fonts.Title
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Location  = New-Object System.Drawing.Point(22, 11)
$lblTitle.AutoSize  = $true
$header.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text      = 'Begeleide upgrade-workflow  |  valideer, dry-run, uitvoeren'
$lblSubtitle.Font      = $script:Fonts.Sub
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(170, 190, 215)
$lblSubtitle.Location  = New-Object System.Drawing.Point(25, 40)
$lblSubtitle.AutoSize  = $true
$header.Controls.Add($lblSubtitle)

$lblHost = New-Object System.Windows.Forms.Label
$lblHost.Text      = "$env:COMPUTERNAME  |  $env:USERNAME  |  PS $($PSVersionTable.PSVersion)"
$lblHost.Font      = $script:Fonts.Sub
$lblHost.ForeColor = [System.Drawing.Color]::FromArgb(150, 172, 200)
$lblHost.AutoSize  = $true
$lblHost.Anchor    = 'Top, Right'
$header.Controls.Add($lblHost)
$header.Add_Resize({ $lblHost.Left = $header.Width - $lblHost.Width - 22; $lblHost.Top = 24 })
$lblHost.Left = 900; $lblHost.Top = 24

# --- Configuratiegebied ----------------------------------------------------
$panelConfig = New-Object System.Windows.Forms.Panel
$panelConfig.Dock      = 'Top'
$panelConfig.Height    = 250
$panelConfig.BackColor = $script:Clr.Bg
$panelConfig.Padding   = New-Object System.Windows.Forms.Padding(16, 12, 16, 6)
$form.Controls.Add($panelConfig)

$cardPaths = New-Card -X 16 -Y 12 -W 880 -H 226 -Anchor 'Top, Left, Right'
$panelConfig.Controls.Add($cardPaths)

New-SectionLabel -Parent $cardPaths -Text 'Configuratie' -X 18 -Y 12 | Out-Null

$script:TxtCustomer  = Add-LabeledTextBox -Parent $cardPaths -Label 'KLANTNAAM' -X 18 -Y 34 -Width 250
$script:TxtAadUpn    = Add-LabeledTextBox -Parent $cardPaths -Label 'ENTRA-ACCOUNT UPN (stap 16, optioneel voorinvullen)' -X 290 -Y 34 -Width 340
$script:TxtLocal     = Add-LabeledTextBox -Parent $cardPaths -Label 'LOCALEXPORTFOLDER (configuratie-exports + backups)' -X 18 -Y 84 -Width 700 -Anchor 'Top, Left, Right'
$script:TxtInstaller = Add-LabeledTextBox -Parent $cardPaths -Label 'INSTALLERPATH' -X 18 -Y 134 -Width 700 -Anchor 'Top, Left, Right'

$btnBrowseLocal     = New-ModernButton -Text 'Bladeren' -X 730 -Y 100 -W 100 -H 25 -Anchor 'Top, Right'
$btnBrowseInstaller = New-ModernButton -Text 'Bladeren' -X 730 -Y 150 -W 100 -H 25 -Anchor 'Top, Right'
$cardPaths.Controls.Add($btnBrowseLocal)
$cardPaths.Controls.Add($btnBrowseInstaller)

$lblDapHint = New-Object System.Windows.Forms.Label
$lblDapHint.Text      = "Stap 1: klantdocumentatie\Microsoft Entra Connect  ->  map 'Upgrade $(Get-Date -Format 'dd-MM-yyyy')'"
$lblDapHint.Font      = $script:Fonts.Caption
$lblDapHint.ForeColor = $script:Clr.Accent
$lblDapHint.Location  = New-Object System.Drawing.Point(18, 184)
$lblDapHint.AutoSize  = $true
$cardPaths.Controls.Add($lblDapHint)

$script:LblSyncPath = New-Object System.Windows.Forms.Label
$script:LblSyncPath.Text      = 'Sync-installatiemap: nog niet gedetecteerd.'
$script:LblSyncPath.Font      = $script:Fonts.Caption
$script:LblSyncPath.ForeColor = $script:Clr.Muted
$script:LblSyncPath.Location  = New-Object System.Drawing.Point(18, 202)
$script:LblSyncPath.AutoSize  = $true
$cardPaths.Controls.Add($script:LblSyncPath)

$cardOptions = New-Card -X 908 -Y 12 -W 470 -H 226 -Anchor 'Top, Right'
$panelConfig.Controls.Add($cardOptions)

New-SectionLabel -Parent $cardOptions -Text 'Opties' -X 18 -Y 12 | Out-Null
$script:ChkDelta       = New-ModernCheckBox -Parent $cardOptions -Text 'Na upgrade automatisch Delta sync starten'    -X 18 -Y 36  -Checked $true
$script:ChkInitial     = New-ModernCheckBox -Parent $cardOptions -Text 'Na upgrade automatisch Initial sync starten'  -X 18 -Y 62  -Checked $false
$script:ChkAutoUpgrade = New-ModernCheckBox -Parent $cardOptions -Text 'AutoUpgrade na afloop op Disabled zetten'     -X 18 -Y 88  -Checked $true
$script:ChkAdminCenter = New-ModernCheckBox -Parent $cardOptions -Text 'M365 Admin Center openen bij stap 18/19'      -X 18 -Y 114 -Checked $true

$btnValidate = New-ModernButton -Text 'Valideer omgeving' -X 18  -Y 180 -W 145 -H 32 -Kind 'Secondary'
$btnWhatIf   = New-ModernButton -Text 'Dry run'           -X 170 -Y 180 -W 110 -H 32 -Kind 'Primary'
$btnExecute  = New-ModernButton -Text 'Uitvoeren'         -X 288 -Y 180 -W 160 -H 32 -Kind 'Danger'
$cardOptions.Controls.Add($btnValidate)
$cardOptions.Controls.Add($btnWhatIf)
$cardOptions.Controls.Add($btnExecute)

# --- Statusbalk ------------------------------------------------------------
$panelStatus = New-Object System.Windows.Forms.Panel
$panelStatus.Dock      = 'Bottom'
$panelStatus.Height    = 34
$panelStatus.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($panelStatus)

$script:LblStatus = New-Object System.Windows.Forms.Label
$script:LblStatus.Text      = 'Gereed.'
$script:LblStatus.Font      = $script:Fonts.Sub
$script:LblStatus.ForeColor = $script:Clr.Muted
$script:LblStatus.Location  = New-Object System.Drawing.Point(18, 9)
$script:LblStatus.AutoSize  = $true
$panelStatus.Controls.Add($script:LblStatus)

$script:Progress = New-Object System.Windows.Forms.ProgressBar
$script:Progress.Location = New-Object System.Drawing.Point(900, 8)
$script:Progress.Size     = New-Object System.Drawing.Size(480, 16)
$script:Progress.Minimum  = 0
$script:Progress.Maximum  = @($script:StepMap.Keys).Count
$script:Progress.Style    = 'Continuous'
$script:Progress.Anchor   = 'Top, Right'
$panelStatus.Controls.Add($script:Progress)

# --- Actiebalk -------------------------------------------------------------
$panelActions = New-Object System.Windows.Forms.Panel
$panelActions.Dock      = 'Bottom'
$panelActions.Height    = 52
$panelActions.BackColor = $script:Clr.Bg
$form.Controls.Add($panelActions)

$btnClear = New-ModernButton -Text 'Leegmaken' -X 16 -Y 12 -W 105 -H 28
$panelActions.Controls.Add($btnClear)

$btnSaveReport = New-ModernButton -Text 'Rapport opslaan' -X 129 -Y 12 -W 130 -H 28
$panelActions.Controls.Add($btnSaveReport)

$lblStartFrom = New-Object System.Windows.Forms.Label
$lblStartFrom.Text      = 'Start vanaf stap:'
$lblStartFrom.Font      = $script:Fonts.Sub
$lblStartFrom.ForeColor = $script:Clr.Muted
$lblStartFrom.Location  = New-Object System.Drawing.Point(280, 18)
$lblStartFrom.AutoSize  = $true
$panelActions.Controls.Add($lblStartFrom)

$script:CmbStartFrom = New-Object System.Windows.Forms.ComboBox
$script:CmbStartFrom.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:CmbStartFrom.FlatStyle     = 'Flat'
$script:CmbStartFrom.Font          = $script:Fonts.Base
$script:CmbStartFrom.Location      = New-Object System.Drawing.Point(383, 14)
$script:CmbStartFrom.Size          = New-Object System.Drawing.Size(340, 26)
$script:CmbStartFrom.Items.Add('Vanaf het begin (alle stappen)') | Out-Null
foreach ($key in $script:StepMap.Keys) { $script:CmbStartFrom.Items.Add($key) | Out-Null }
$script:CmbStartFrom.SelectedIndex = 0
$panelActions.Controls.Add($script:CmbStartFrom)

$script:ChkOnlyIssues = New-ModernCheckBox -Parent $panelActions -Text 'Toon alleen aandachtspunten' -X 745 -Y 18 -Checked $false

# --- Split: resultaten + log -----------------------------------------------
$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock             = 'Fill'
$split.Orientation      = 'Horizontal'
$split.BackColor        = $script:Clr.Bg
$split.SplitterWidth    = 8
$split.Panel1.Padding   = New-Object System.Windows.Forms.Padding(16, 6, 16, 0)
$split.Panel2.Padding   = New-Object System.Windows.Forms.Padding(16, 0, 16, 6)
$form.Controls.Add($split)

$script:Grid = New-Object System.Windows.Forms.DataGridView
$script:Grid.Dock                     = 'Fill'
$script:Grid.AllowUserToAddRows       = $false
$script:Grid.AllowUserToDeleteRows    = $false
$script:Grid.AllowUserToResizeRows    = $false
$script:Grid.ReadOnly                 = $true
$script:Grid.RowHeadersVisible        = $false
$script:Grid.SelectionMode            = 'FullRowSelect'
$script:Grid.AutoSizeColumnsMode      = 'Fill'
$script:Grid.BorderStyle              = 'None'
$script:Grid.CellBorderStyle          = 'SingleHorizontal'
$script:Grid.BackgroundColor          = [System.Drawing.Color]::White
$script:Grid.GridColor                = $script:Clr.Border
$script:Grid.EnableHeadersVisualStyles = $false
$script:Grid.ColumnHeadersHeight      = 34
$script:Grid.ColumnHeadersHeightSizeMode = 'DisableResizing'
$script:Grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)
$script:Grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:Clr.Muted
$script:Grid.ColumnHeadersDefaultCellStyle.Font      = $script:Fonts.Bold
$script:Grid.ColumnHeadersDefaultCellStyle.Padding   = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$script:Grid.DefaultCellStyle.Font                   = $script:Fonts.Base
$script:Grid.DefaultCellStyle.ForeColor              = $script:Clr.Text
$script:Grid.DefaultCellStyle.SelectionBackColor     = $script:Clr.AccentLight
$script:Grid.DefaultCellStyle.SelectionForeColor     = $script:Clr.Text
$script:Grid.DefaultCellStyle.Padding                = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$script:Grid.AlternatingRowsDefaultCellStyle.BackColor = $script:Clr.RowAlt
$script:Grid.RowTemplate.Height       = 28
$script:Grid.ScrollBars               = 'Vertical'

$null = $script:Grid.Columns.Add('Time',   'Tijd')
$null = $script:Grid.Columns.Add('Step',   'Stap')
$null = $script:Grid.Columns.Add('Title',  'Titel')
$null = $script:Grid.Columns.Add('Status', 'Status')
$null = $script:Grid.Columns.Add('Detail', 'Detail')
$script:Grid.Columns[0].FillWeight = 7
$script:Grid.Columns[1].FillWeight = 9
$script:Grid.Columns[2].FillWeight = 25
$script:Grid.Columns[3].FillWeight = 10
$script:Grid.Columns[4].FillWeight = 49
$split.Panel1.Controls.Add($script:Grid)

$script:TxtLog = New-Object System.Windows.Forms.RichTextBox
$script:TxtLog.Dock        = 'Fill'
$script:TxtLog.ReadOnly    = $true
$script:TxtLog.BorderStyle = 'None'
$script:TxtLog.BackColor   = $script:Clr.LogBg
$script:TxtLog.ForeColor   = $script:Clr.LogText
$script:TxtLog.Font        = $script:Fonts.Mono
$script:TxtLog.WordWrap    = $true
$split.Panel2.Controls.Add($script:TxtLog)

# --- Dock-volgorde ---------------------------------------------------------
# WinForms dockt van de achterste naar de voorste control (hoogste index in
# Controls eerst). Wordt de Fill-control als laatste toegevoegd, dan wordt die
# als EERSTE gedockt en valt het grid deels achter de header/config- en
# actiepanelen. Daarom wordt de volgorde hier expliciet gezet:
#   statusbalk (onderste rand) -> actiebalk -> header (bovenrand) -> config -> split (rest)
$form.SuspendLayout()
$split.BringToFront()
$panelConfig.SendToBack()
$header.SendToBack()
$panelActions.SendToBack()
$panelStatus.SendToBack()
$form.ResumeLayout($true)

$script:ActionControls = @($btnValidate, $btnWhatIf, $btnExecute, $btnClear, $btnSaveReport, $script:CmbStartFrom)

$form.Add_Load({
    $available = $split.Height
    if ($available -gt 200) { $split.SplitterDistance = [Math]::Max(180, [int]($available * 0.58)) }
    $lblHost.Left = $header.Width - $lblHost.Width - 22
})
$form.Add_Resize({
    $available = $split.Height
    if ($available -gt 200) { $split.SplitterDistance = [Math]::Max(180, [int]($available * 0.58)) }
})

# --- Dialogen en events ----------------------------------------------------
$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$openDialog   = New-Object System.Windows.Forms.OpenFileDialog
$openDialog.Filter = 'Installers (*.exe;*.msi)|*.exe;*.msi|Alle bestanden (*.*)|*.*'

$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($btnValidate, 'Controleert de omgeving zonder iets te wijzigen.')
$tip.SetToolTip($btnWhatIf,   'Loopt de volledige workflow door zonder wijzigingen door te voeren.')
$tip.SetToolTip($btnExecute,  'Voert echte wijzigingen uit op deze Entra Connect server.')
$tip.SetToolTip($script:CmbStartFrom, 'Hervat de workflow vanaf een specifieke stap na een onderbreking.')

$btnBrowseLocal.Add_Click({ if ($folderDialog.ShowDialog() -eq 'OK') { $script:TxtLocal.Text = $folderDialog.SelectedPath } })
$btnBrowseInstaller.Add_Click({ if ($openDialog.ShowDialog() -eq 'OK') { $script:TxtInstaller.Text = $openDialog.FileName } })

$script:CmbStartFrom.Add_SelectedIndexChanged({
    if ($script:CmbStartFrom.SelectedIndex -le 0) {
        $script:StartFromStep = $null
    } else {
        $script:StartFromStep = $script:StepMap[$script:CmbStartFrom.SelectedItem.ToString()]
    }
})

$script:ChkOnlyIssues.Add_CheckedChanged({ Update-GridFilter })

$script:Grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $row = $sender.Rows[$e.RowIndex]
    if ($null -eq $row.Tag) { return }
    $r = $row.Tag
    $text = "Stap   : $($r.Step)`r`nTitel  : $($r.Title)`r`nStatus : $($r.Status)`r`nTijd   : $($r.Timestamp)`r`n`r`n$($r.Detail)"
    [System.Windows.Forms.MessageBox]::Show($text, 'Detail', 'OK', 'Information') | Out-Null
})

$btnClear.Add_Click({ Clear-Results })

$btnSaveReport.Add_Click({
    $cfg = Get-InputState
    $target = if (-not [string]::IsNullOrWhiteSpace($cfg.LocalExportFolder)) { $cfg.LocalExportFolder } else { $env:TEMP }
    Save-ResultsReport -Folder $target
    [System.Windows.Forms.MessageBox]::Show("Rapport opgeslagen in: $target", 'Rapport', 'OK', 'Information') | Out-Null
})

$btnValidate.Add_Click({
    Clear-Results
    Set-StatusText 'Validatie bezig...'
    Write-Log 'Validatie van de omgeving gestart.'
    Set-UiBusy $true
    try {
        Validate-CoreEnvironment
    } catch {
        New-Result 'FOUT' 'Onverwachte fout tijdens validatie' 'BLOKKADE' $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Fout', 'OK', 'Error') | Out-Null
    } finally {
        Set-UiBusy $false
    }
    Set-StatusText 'Validatie afgerond.'
})

$btnWhatIf.Add_Click({
    Clear-Results
    Set-StatusText 'Dry run bezig...'
    Write-Log 'Dry run gestart.'
    Set-UiBusy $true
    try {
        Invoke-WhatIfWorkflow
        [System.Windows.Forms.MessageBox]::Show('Dry run afgerond. Controleer de resultaten en het logvenster.', 'Klaar', 'OK', 'Information') | Out-Null
    } catch {
        New-Result 'FOUT' 'Onverwachte fout in dry run' 'BLOKKADE' $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Fout', 'OK', 'Error') | Out-Null
    } finally {
        Set-UiBusy $false
    }
    Set-StatusText 'Dry run afgerond.'
})

$btnExecute.Add_Click({
    $startMsg = if ([string]::IsNullOrWhiteSpace($script:StartFromStep)) { 'alle stappen (vanaf het begin)' } else { "vanaf $script:StartFromStep" }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Dit voert echte wijzigingen uit op de Entra Connect server.`r`nWorkflow start: $startMsg`r`n`r`nDoorgaan?",
        'Bevestiging', 'OKCancel', 'Warning')
    if ($answer -ne [System.Windows.Forms.DialogResult]::OK) { return }

    Clear-Results
    Set-StatusText 'Workflow bezig...'
    Write-Log "Workflow gestart ($startMsg)." -Level 'Warn'
    Set-UiBusy $true
    $script:ServiceStoppedByScript = $false
    try {
        Invoke-ExecuteWorkflow
        Restore-ADSyncServiceIfNeeded
        [System.Windows.Forms.MessageBox]::Show('Workflow afgerond. Controleer de resultaten en het logvenster.', 'Klaar', 'OK', 'Information') | Out-Null
    } catch {
        New-Result 'FOUT' 'Onverwachte fout' 'BLOKKADE' $_.Exception.Message
        Restore-ADSyncServiceIfNeeded
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Fout', 'OK', 'Error') | Out-Null
    } finally {
        Set-UiBusy $false
    }
    Set-StatusText 'Workflow afgerond.'
})

Write-Log 'GUI geladen. Vul de paden in en start met Valideer omgeving of Dry run.'
Update-SyncPathLabel | Out-Null
if (-not (Test-IsAdministrator)) {
    Write-Log 'Let op: deze sessie draait NIET als Administrator. Uitvoeren zal falen.' -Level 'Warn'
}

[void]$form.ShowDialog()
