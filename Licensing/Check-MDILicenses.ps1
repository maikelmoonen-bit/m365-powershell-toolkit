#Requires -Version 5.1
<#
.SYNOPSIS
    Inventariseert gebruikerslicenties voor Microsoft Defender for Identity (MDI).

.DESCRIPTION
    Dit script haalt alle gebruikers op uit Microsoft Entra ID (Azure AD) via
    Microsoft Graph, filtert op gebruikers met een ingevuld EmployeeID (= in scope),
    en beoordeelt of zij beschikken over een van de vereiste licenties voor MDI.

    Output:
      - Consolerapport met samenvatting
      - Visueel opgemaakte Excel (.xlsx) met:
          • Tabblad "Dashboard"     – KPI-overzichtskaarten + taartdiagram
          • Tabblad "Alle gebruikers"    – volledig gekleurd overzicht
          • Tabblad "Aanschaf vereist"   – gefilterd, rood gemarkeerd
          • Tabblad "Licentie aanwezig"  – gefilterd, groen gemarkeerd

.PARAMETER OutputPath
    Pad naar de output-map. Standaard: huidige werkmap.

.PARAMETER ExportFileName
    Bestandsnaam voor de Excel-export. Standaard: MDI_License_Check_<datum>.xlsx

.PARAMETER TenantId
    Optioneel: Tenant ID om in te loggen. Als weggelaten wordt de standaard tenant gebruikt.

.EXAMPLE
    .\Check-MDILicenses.ps1

.EXAMPLE
    .\Check-MDILicenses.ps1 -OutputPath "C:\Reports" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.NOTES
    Vereisten:
      - PowerShell 5.1 of hoger (PowerShell 7+ aanbevolen)
      - Microsoft.Graph PowerShell module   (Install-Module Microsoft.Graph)
      - ImportExcel PowerShell module        (Install-Module ImportExcel)
      - Minimale Graph-rechten: User.Read.All, Organization.Read.All
        (of Directory.Read.All voor vereenvoudigd beheer)

    Auteur  : Maikel Moonen - Moonen IT Consulting
    Versie  : 2.1 - toegevoegd: generieke detectie via Service Plan ID, waardoor
              "Microsoft Defender Suite for Microsoft 365 Business Premium" en
              andere (toekomstige) bundels die MDI bevatten automatisch als
              kwalificerend worden herkend, naast de vaste SKU-lijst.
    Datum   : 2025
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter()]
    [string]$ExportFileName = "MDI_License_Check_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",

    [Parameter()]
    [string]$TenantId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─────────────────────────────────────────────────────────────────────────────
# SECTIE 1 – SKU-mapping: vriendelijke naam → Service Plan ID's / SKU Part Numbers
# ─────────────────────────────────────────────────────────────────────────────
# Elke licentie wordt herkend via de officiële SKU PartNumber (zoals teruggegeven
# door Microsoft Graph). Aliassen dekken varianten (A5, E5, G5, F5 etc.).
#
# Bronnen:
#   https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
# ─────────────────────────────────────────────────────────────────────────────

$MDI_QUALIFYING_SKUS = @(
    # Enterprise Mobility + Security E5 / A5
    "EMSPREMIUM",           # EMS E5
    "EMSPREMIUM_FACULTY",   # EMS A5 Faculty
    "EMSPREMIUM_STUDENT",   # EMS A5 Student

    # Microsoft 365 E5
    "SPE_E5",               # Microsoft 365 E5
    "SPE_E5_NOPSTNCONF",    # Microsoft 365 E5 (no audio conferencing)

    # Microsoft 365 A5
    "M365EDU_A5_FACULTY",   # Microsoft 365 A5 Faculty
    "M365EDU_A5_STUDENT",   # Microsoft 365 A5 Student

    # Microsoft 365 G5 (Government)
    "M365_G5",              # Microsoft 365 G5
    "M365_G5_GOV",

    # Microsoft 365 F5 Security
    "M365_F5_SECURITY",     # Microsoft 365 F5 Security
    "M365_F5_SECURITY_COMPLIANCE", # Microsoft 365 F5 Security + Compliance

    # Microsoft 365 Business Premium
    "SPB",                  # Microsoft 365 Business Premium (SkuPartNumber)

    # Microsoft Defender for Office 365 Plan 1 (standalone)
    "DEFENDER_ENDPOINT_P1", # MDE P1 standalone – soms ook via ATP-bundel
    "ATP_ENTERPRISE",       # Microsoft Defender for Office 365 Plan 1 (oud)
    "MDATP",                # Microsoft Defender for Endpoint Plan 1 (alternatief)

    # Microsoft 365 A5/E5/G5/F5 Security (gebundeld)
    "IDENTITY_THREAT_PROTECTION",       # M365 E5 Security add-on
    "IDENTITY_THREAT_PROTECTION_FOR_EMS_E5"
)

# ─────────────────────────────────────────────────────────────────────────────
# SECTIE 1b – Generieke detectie via Service Plan ID (aanvullend op de lijst hierboven)
# ─────────────────────────────────────────────────────────────────────────────
# In plaats van alleen te vertrouwen op een vaste lijst met SKU-namen (die Microsoft
# regelmatig wijzigt/uitbreidt - zoals recent met "Microsoft Defender Suite for
# Microsoft 365 Business Premium"), wordt elke licentie die het Service Plan
# "Microsoft Defender for Identity" bevat automatisch als kwalificerend herkend.
# Dit Service Plan ID is stabiel, onafhankelijk van hoe Microsoft de SKU zelf noemt.
#
# Bron: https://learn.microsoft.com/entra/identity/users/licensing-service-plan-reference
#   Service plan "Microsoft Defender for Identity" (String ID: ATA)
#   Service Plan ID: 14ab5db5-e6c4-4b20-b4bc-13e36fd2227f
#
# Hierdoor worden zowel bekende SKU's (EMS E5, M365 E5, etc.) als nieuwe/onbekende
# bundels die dit onderdeel bevatten - zoals "Microsoft Defender Suite for
# Microsoft 365 Business Premium" (sinds sep. 2025, nog zonder vaste SkuPartNumber
# in de officiele Microsoft-referentie) - automatisch als kwalificerend gezien,
# zonder dat de vaste lijst hierboven steeds handmatig bijgewerkt moet worden.
$MDI_SERVICE_PLAN_ID = "14ab5db5-e6c4-4b20-b4bc-13e36fd2227f"

# ─────────────────────────────────────────────────────────────────────────────
# SECTIE 2 – Vriendelijke weergavenamen per SKU (voor leesbare output)
# ─────────────────────────────────────────────────────────────────────────────
$SKU_DISPLAY_NAMES = @{
    "EMSPREMIUM"                              = "Enterprise Mobility + Security E5"
    "EMSPREMIUM_FACULTY"                      = "Enterprise Mobility + Security A5 (Faculty)"
    "EMSPREMIUM_STUDENT"                      = "Enterprise Mobility + Security A5 (Student)"
    "SPE_E5"                                  = "Microsoft 365 E5"
    "SPE_E5_NOPSTNCONF"                       = "Microsoft 365 E5 (No Audio Conferencing)"
    "M365EDU_A5_FACULTY"                      = "Microsoft 365 A5 (Faculty)"
    "M365EDU_A5_STUDENT"                      = "Microsoft 365 A5 (Student)"
    "M365_G5"                                 = "Microsoft 365 G5"
    "M365_G5_GOV"                             = "Microsoft 365 G5 (Government)"
    "M365_F5_SECURITY"                        = "Microsoft 365 F5 Security"
    "M365_F5_SECURITY_COMPLIANCE"             = "Microsoft 365 F5 Security + Compliance"
    "SPB"                                     = "Microsoft 365 Business Premium"
    "DEFENDER_ENDPOINT_P1"                    = "Microsoft Defender for Endpoint Plan 1"
    "ATP_ENTERPRISE"                          = "Microsoft Defender for Office 365 Plan 1"
    "MDATP"                                   = "Microsoft Defender for Endpoint"
    "IDENTITY_THREAT_PROTECTION"              = "Microsoft 365 E5 Security"
    "IDENTITY_THREAT_PROTECTION_FOR_EMS_E5"   = "Microsoft 365 E5 Security (EMS)"
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTIE 3 – Hulpfuncties
# ─────────────────────────────────────────────────────────────────────────────

function Write-Header {
    param([string]$Title)
    $line = "=" * 70
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
}

function Write-Step {
    param([string]$Message)
    Write-Host "[*] $Message" -ForegroundColor Yellow
}

function Write-OK {
    param([string]$Message)
    Write-Host "[+] $Message" -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)
    Write-Host "[!] $Message" -ForegroundColor Red
}

function Ensure-GraphModule {
    Write-Step "Controleren Microsoft.Graph module..."
    if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Users")) {
        Write-Warn "Microsoft.Graph module niet gevonden. Installeren..."
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Write-OK "Microsoft.Graph module geladen."

    Write-Step "Controleren ImportExcel module..."
    if (-not (Get-Module -ListAvailable -Name "ImportExcel")) {
        Write-Warn "ImportExcel module niet gevonden. Installeren..."
        Install-Module ImportExcel -Scope CurrentUser -Force
    }
    Import-Module ImportExcel -ErrorAction Stop
    Write-OK "ImportExcel module geladen."
}

function Connect-ToGraph {
    param([string]$TenantId)
    Write-Step "Verbinden met Microsoft Graph..."

    $scopes = @("User.Read.All", "Organization.Read.All")

    $connectParams = @{ Scopes = $scopes }
    if ($TenantId) { $connectParams["TenantId"] = $TenantId }

    Connect-MgGraph @connectParams -ErrorAction Stop
    Write-OK "Verbonden met Microsoft Graph."
}

function Get-AllUsersWithLicenses {
    <#
    .SYNOPSIS
        Haalt alle gebruikers op, inclusief hun licentie-informatie en EmployeeId.
    #>
    Write-Step "Gebruikers ophalen uit Entra ID (dit kan even duren)..."

    # Selecteer alleen de benodigde velden voor efficiëntie
    $selectFields = @(
        "id",
        "displayName",
        "userPrincipalName",
        "mail",
        "employeeId",
        "accountEnabled",
        "userType",
        "assignedLicenses",
        "department",
        "jobTitle"
    )

    $users = Get-MgUser -All `
        -Select ($selectFields -join ",") `
        -Filter "assignedLicenses/`$count ne 0" `
        -ConsistencyLevel eventual `
        -CountVariable userCount

    Write-OK "$($users.Count) gebruikers gevonden met ten minste één licentie."
    return $users
}

function Get-TenantSkuMap {
    <#
    .SYNOPSIS
        Haalt de beschikbare SKU's op in de tenant en maakt een lookup-tabel
        van SkuId (GUID) naar SKU-details (PartNumber + bevatten Service Plan ID's).
    #>
    Write-Step "Tenant licentie-SKU's ophalen..."
    $skus = Get-MgSubscribedSku -All
    $map  = @{}
    foreach ($sku in $skus) {
        $servicePlanIds = @($sku.ServicePlans | ForEach-Object { $_.ServicePlanId })
        $map[$sku.SkuId] = [PSCustomObject]@{
            PartNumber     = $sku.SkuPartNumber
            ServicePlanIds = $servicePlanIds
        }
    }
    Write-OK "$($map.Count) SKU-definities opgehaald uit tenant."
    return $map
}

function Test-UserHasMDILicense {
    <#
    .SYNOPSIS
        Controleert of een gebruiker minimaal één kwalificerende MDI-licentie heeft.
    .DESCRIPTION
        Een licentie wordt als kwalificerend beschouwd als:
          1) De SkuPartNumber voorkomt in de vaste lijst $MDI_QUALIFYING_SKUS, OF
          2) De SKU het Service Plan "Microsoft Defender for Identity"
             ($MDI_SERVICE_PLAN_ID) bevat EN dat onderdeel niet specifiek is
             uitgeschakeld (DisabledPlans) voor deze gebruiker.
        Methode 2 is generiek en naam-onafhankelijk: hierdoor wordt o.a. "Microsoft
        Defender Suite for Microsoft 365 Business Premium" - en elke andere huidige
        of toekomstige bundel die MDI bevat - automatisch herkend, ook zonder dat de
        exacte SkuPartNumber van die bundel bekend of in de lijst opgenomen is.
    .RETURNS
        Hashtable met:
          HasQualifyingLicense [bool]
          QualifyingLicenses   [string[]]  – vriendelijke namen
          AllLicenses          [string[]]  – alle toegewezen SKU's (PartNumbers)
    #>
    param(
        [object]$User,
        [hashtable]$SkuMap
    )

    $allPartNumbers      = @()
    $qualifyingLicenses  = @()

    foreach ($assignedLicense in $User.AssignedLicenses) {
        $skuId         = $assignedLicense.SkuId
        $disabledPlans = @($assignedLicense.DisabledPlans)
        $skuInfo       = $SkuMap[$skuId]

        if ($skuInfo) {
            $partNumber = $skuInfo.PartNumber
            $allPartNumbers += $partNumber

            # Methode 1: SKU staat op de vaste lijst met bekende kwalificerende licenties
            $isQualifyingBySkuList = ($MDI_QUALIFYING_SKUS -contains $partNumber)

            # Methode 2: SKU bevat het Service Plan "Microsoft Defender for Identity"
            # en dat onderdeel is niet apart uitgeschakeld voor deze gebruiker
            $isQualifyingByServicePlan = ($skuInfo.ServicePlanIds -contains $MDI_SERVICE_PLAN_ID) `
                                          -and ($disabledPlans -notcontains $MDI_SERVICE_PLAN_ID)

            if ($isQualifyingBySkuList -or $isQualifyingByServicePlan) {
                $friendlyName = if ($SKU_DISPLAY_NAMES.ContainsKey($partNumber)) {
                    $SKU_DISPLAY_NAMES[$partNumber]
                } elseif ($isQualifyingByServicePlan) {
                    "$partNumber (bevat Microsoft Defender for Identity)"
                } else {
                    $partNumber
                }
                $qualifyingLicenses += $friendlyName
            }
        } else {
            # Onbekende SKU-GUID (bijv. preview-licentie)
            $allPartNumbers += "Onbekend ($skuId)"
        }
    }

    return @{
        HasQualifyingLicense = ($qualifyingLicenses.Count -gt 0)
        QualifyingLicenses   = $qualifyingLicenses
        AllLicenses          = $allPartNumbers
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SECTIE 4 – Hoofdlogica
# ─────────────────────────────────────────────────────────────────────────────

function Main {
    Write-Header "Microsoft Defender for Identity – Licentie Inventarisatie"

    # Stap 1: Module & connectie
    Ensure-GraphModule
    Connect-ToGraph -TenantId $TenantId

    # Stap 2: Data ophalen
    $allLicensedUsers = Get-AllUsersWithLicenses
    $skuMap           = Get-TenantSkuMap

    # Stap 3: Filteren op EmployeeId (= in-scope gebruikers)
    Write-Step "Filteren op gebruikers met ingevuld EmployeeID..."
    $inScopeUsers = $allLicensedUsers | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.EmployeeId)
    }
    Write-OK "$($inScopeUsers.Count) gebruikers in scope (EmployeeID ingevuld)."

    if ($inScopeUsers.Count -eq 0) {
        Write-Warn "Geen in-scope gebruikers gevonden. Script wordt beëindigd."
        Disconnect-MgGraph | Out-Null
        return
    }

    # Stap 4: Licentiecontrole per gebruiker
    Write-Step "Licenties controleren per gebruiker..."

    $results          = [System.Collections.Generic.List[PSCustomObject]]::new()
    $needsLicense     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $hasLicense       = [System.Collections.Generic.List[PSCustomObject]]::new()

    $counter = 0
    foreach ($user in $inScopeUsers) {
        $counter++
        Write-Progress -Activity "Licenties controleren" `
                       -Status "$counter / $($inScopeUsers.Count): $($user.UserPrincipalName)" `
                       -PercentComplete (($counter / $inScopeUsers.Count) * 100)

        $licenseCheck = Test-UserHasMDILicense -User $user -SkuMap $skuMap

        $status = if ($licenseCheck.HasQualifyingLicense) {
            "LICENTIE AANWEZIG – Geen actie nodig"
        } else {
            "LICENTIE ONTBREEKT – Aanschaf vereist"
        }

        $resultObj = [PSCustomObject]@{
            Status                  = $status
            DisplayName             = $user.DisplayName
            UserPrincipalName       = $user.UserPrincipalName
            Mail                    = $user.Mail
            EmployeeID              = $user.EmployeeId
            Department              = $user.Department
            JobTitle                = $user.JobTitle
            AccountEnabled          = $user.AccountEnabled
            HeeftKwalificerndeLic   = $licenseCheck.HasQualifyingLicense
            KwalificerndeLicenties  = ($licenseCheck.QualifyingLicenses -join " | ")
            AlleLicenties           = ($licenseCheck.AllLicenses -join " | ")
        }

        $results.Add($resultObj)

        if ($licenseCheck.HasQualifyingLicense) {
            $hasLicense.Add($resultObj)
        } else {
            $needsLicense.Add($resultObj)
        }
    }

    Write-Progress -Activity "Licenties controleren" -Completed

    # Stap 5: Samenvatting tonen
    Write-Header "Samenvatting"
    Write-Host ""
    Write-Host "  Totaal in-scope gebruikers (EmployeeID ingevuld) : $($inScopeUsers.Count)"
    Write-OK   "  Gebruikers met kwalificerende MDI-licentie        : $($hasLicense.Count)"
    Write-Warn "  Gebruikers ZONDER kwalificerende licentie         : $($needsLicense.Count)"
    Write-Host ""

    if ($needsLicense.Count -gt 0) {
        Write-Host "  ─── Gebruikers waarvoor licenties aangeschaft moeten worden ───" -ForegroundColor Red
        $needsLicense | Format-Table DisplayName, UserPrincipalName, Department, AlleLicenties -AutoSize
    }

    # Stap 6: Exporteren naar Excel (.xlsx)
    $exportFilePath = Join-Path -Path $OutputPath -ChildPath $ExportFileName
    Write-Step "Excel-rapport genereren: $exportFilePath"

    # Verwijder eventueel bestaand bestand (ImportExcel kan daar soms mee conflicteren)
    if (Test-Path $exportFilePath) { Remove-Item $exportFilePath -Force }

    # ── Kleurpalet ──────────────────────────────────────────────────────────
    $colorBlauw       = "2E5FA3"   # Microsoft-blauw – kopteksten
    $colorBlauwLicht  = "D6E4F0"   # Lichtblauw – alternatieve rijen
    $colorRood        = "C0392B"   # Rood – ontbrekende licentie
    $colorRoodLicht   = "FADBD8"   # Lichtroze – rij zonder licentie
    $colorGroen       = "1E8449"   # Groen – licentie aanwezig
    $colorGroenLicht  = "D5F5E3"   # Lichtgroen – rij met licentie
    $colorGrijs       = "F2F3F4"   # Lichtgrijs – Dashboard achtergrond
    $colorWit         = "FFFFFF"
    $colorTekstWit    = "FFFFFF"
    $colorTekstDonker = "1A1A2E"

    # ── Helpersfunctie: kolombreedtes instellen ──────────────────────────────
    function Set-ColumnWidths {
        param($Worksheet, [hashtable]$Widths)
        foreach ($col in $Widths.Keys) {
            $Worksheet.Column($col).Width = $Widths[$col]
        }
    }

    # ── Helpersfunctie: stijl op bereik toepassen ───────────────────────────
    function Set-RangeStyle {
        param($Worksheet, [string]$Range,
              [string]$BgColor    = "",
              [string]$FontColor  = "",
              [bool]$Bold         = $false,
              [int]$FontSize      = 0,
              [string]$HAlign     = "",
              [bool]$WrapText     = $false,
              [bool]$Border       = $false)

        $r = $Worksheet.Cells[$Range]
        if ($BgColor)   { $r.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                          $r.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$BgColor")) }
        if ($FontColor) { $r.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$FontColor")) }
        if ($Bold)      { $r.Style.Font.Bold = $true }
        if ($FontSize)  { $r.Style.Font.Size = $FontSize }
        if ($HAlign)    { $r.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::$HAlign }
        if ($WrapText)  { $r.Style.WrapText = $true }
        if ($Border) {
            foreach ($side in @("Top","Bottom","Left","Right")) {
                $r.Style.Border.$side.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                $r.Style.Border.$side.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#BDBDBD"))
            }
        }
    }

    # ════════════════════════════════════════════════════════════════════════
    # TABBLAD 1 – DASHBOARD
    # ════════════════════════════════════════════════════════════════════════
    Write-Step "  Tabblad 1: Dashboard opbouwen..."

    $excelPkg = [OfficeOpenXml.ExcelPackage]::new([System.IO.FileInfo]$exportFilePath)
    $wsDash   = $excelPkg.Workbook.Worksheets.Add("Dashboard")

    # Achtergrond heel werkblad grijs
    Set-RangeStyle $wsDash "A1:Z60" -BgColor $colorGrijs

    # ── Titelbalk ──
    $wsDash.Cells["A1:J3"].Merge = $true
    $wsDash.Cells["A1"].Value    = "Microsoft Defender for Identity – Licentie Inventarisatie"
    Set-RangeStyle $wsDash "A1:J3" -BgColor $colorBlauw -FontColor $colorTekstWit -Bold $true -FontSize 18 -HAlign "Center"
    $wsDash.Row(1).Height = 45

    # Subtitel / datum
    $wsDash.Cells["A4:J4"].Merge = $true
    $wsDash.Cells["A4"].Value    = "Rapportdatum: $(Get-Date -Format 'dd-MM-yyyy HH:mm')   |   Gegenereerd door: Check-MDILicenses.ps1 v2.1"
    Set-RangeStyle $wsDash "A4:J4" -BgColor $colorBlauw -FontColor "B3C6E0" -FontSize 10 -HAlign "Center"
    $wsDash.Row(4).Height = 20

    # ── KPI-kaarten (rij 6-9) ──
    # Kaart 1: Totaal in scope
    $wsDash.Cells["B6:D9"].Merge = $true
    $wsDash.Cells["B6"].Value    = $inScopeUsers.Count
    Set-RangeStyle $wsDash "B6:D9" -BgColor $colorBlauw -FontColor $colorTekstWit -Bold $true -FontSize 36 -HAlign "Center"
    $wsDash.Cells["B10:D10"].Merge = $true
    $wsDash.Cells["B10"].Value   = "TOTAAL IN SCOPE"
    Set-RangeStyle $wsDash "B10:D10" -BgColor "1A3A6B" -FontColor $colorTekstWit -Bold $true -FontSize 10 -HAlign "Center"

    # Kaart 2: Licentie aanwezig
    $wsDash.Cells["F6:H9"].Merge = $true
    $wsDash.Cells["F6"].Value    = $hasLicense.Count
    Set-RangeStyle $wsDash "F6:H9" -BgColor $colorGroen -FontColor $colorTekstWit -Bold $true -FontSize 36 -HAlign "Center"
    $wsDash.Cells["F10:H10"].Merge = $true
    $wsDash.Cells["F10"].Value   = "LICENTIE AANWEZIG ✔"
    Set-RangeStyle $wsDash "F10:H10" -BgColor "145A32" -FontColor $colorTekstWit -Bold $true -FontSize 10 -HAlign "Center"

    # Kaart 3: Aanschaf vereist
    $wsDash.Cells["J6:L9"].Merge = $true
    $wsDash.Cells["J6"].Value    = $needsLicense.Count
    $kpiKleur = if ($needsLicense.Count -gt 0) { $colorRood } else { $colorGroen }
    Set-RangeStyle $wsDash "J6:L9" -BgColor $kpiKleur -FontColor $colorTekstWit -Bold $true -FontSize 36 -HAlign "Center"
    $wsDash.Cells["J10:L10"].Merge = $true
    $wsDash.Cells["J10"].Value   = "AANSCHAF VEREIST ✖"
    $kpiDarkKleur = if ($needsLicense.Count -gt 0) { "922B21" } else { "145A32" }
    Set-RangeStyle $wsDash "J10:L10" -BgColor $kpiDarkKleur -FontColor $colorTekstWit -Bold $true -FontSize 10 -HAlign "Center"

    foreach ($row in 6..10) { $wsDash.Row($row).Height = 22 }

    # ── Taartdiagram ──
    $chartSheet = $wsDash.Drawings.AddChart("LicentieVerdeling", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie3D)
    $chartSheet.Title.Text      = "Licentieverdeling in scope"
    $chartSheet.SetPosition(11, 0, 1, 0)   # rij 12, kolom B
    $chartSheet.SetSize(480, 280)

    # Gegevens voor chart in verborgen helper-cellen (kolom P)
    $wsDash.Cells["P1"].Value = "Categorie"
    $wsDash.Cells["P2"].Value = "Licentie aanwezig"
    $wsDash.Cells["P3"].Value = "Aanschaf vereist"
    $wsDash.Cells["Q1"].Value = "Aantal"
    $wsDash.Cells["Q2"].Value = $hasLicense.Count
    $wsDash.Cells["Q3"].Value = $needsLicense.Count

    $series = $chartSheet.Series.Add($wsDash.Cells["Q2:Q3"], $wsDash.Cells["P2:P3"])
    $series.Header = "Gebruikers"

    # ── Legenda kwalificerende licenties ──
    $wsDash.Cells["B30:L30"].Merge = $true
    $wsDash.Cells["B30"].Value     = "Kwalificerende licenties voor Microsoft Defender for Identity"
    Set-RangeStyle $wsDash "B30:L30" -BgColor $colorBlauw -FontColor $colorTekstWit -Bold $true -FontSize 11 -HAlign "Left"
    $wsDash.Row(30).Height = 18

    $licList = @(
        "Enterprise Mobility + Security E5/A5",
        "Microsoft 365 A5 / E5 / G5",
        "Microsoft 365 A5 / E5 / G5 / F5 Security",
        "Microsoft 365 Business Premium",
        "Microsoft 365 F5 Security + Compliance",
        "Microsoft Defender for Office 365 Plan 1 (standalone)",
        "Microsoft Defender Suite for Microsoft 365 Business Premium",
        "+ elke andere licentie die Microsoft Defender for Identity bevat (auto-detectie)"
    )
    $licRow = 31
    foreach ($lic in $licList) {
        $wsDash.Cells["B$licRow"].Value = "  ✔  $lic"
        $bgC = if ($licRow % 2 -eq 0) { $colorBlauwLicht } else { $colorWit }
        Set-RangeStyle $wsDash "B${licRow}:L${licRow}" -BgColor $bgC -FontSize 10
        $wsDash.Cells["B$licRow"].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$colorGroen"))
        $wsDash.Row($licRow).Height = 16
        $licRow++
    }

    # Kolombreedtes Dashboard
    Set-ColumnWidths $wsDash @{ 1=3; 2=12; 3=12; 4=12; 5=3; 6=12; 7=12; 8=12; 9=3; 10=12; 11=12; 12=12 }

    # ════════════════════════════════════════════════════════════════════════
    # HULPFUNCTIE: Voeg een data-tabblad toe
    # ════════════════════════════════════════════════════════════════════════
    function Add-DataSheet {
        param(
            [OfficeOpenXml.ExcelPackage]$Pkg,
            [string]$SheetName,
            [object[]]$Data,
            [string]$HeaderBg,
            [string]$AltRowBg,
            [string]$StatusBgTrue,
            [string]$StatusBgFalse
        )

        $ws = $Pkg.Workbook.Worksheets.Add($SheetName)

        # Bevroren rij + AutoFilter
        $ws.View.FreezePanes(2, 1)

        if ($Data.Count -eq 0) {
            $ws.Cells["A1"].Value = "Geen gebruikers in deze categorie."
            Set-RangeStyle $ws "A1:K1" -BgColor $HeaderBg -FontColor $colorTekstWit -Bold $true
            $Pkg.Workbook.Worksheets[$SheetName] | Out-Null
            return $ws
        }

        # Kolomkoppen
        $headers = @(
            "Status", "Naam", "UserPrincipalName", "E-mailadres",
            "EmployeeID", "Afdeling", "Functie",
            "Account actief", "MDI licentie aanwezig",
            "Kwalificerende licentie(s)", "Alle toegewezen licenties"
        )

        for ($c = 0; $c -lt $headers.Count; $c++) {
            $cell = $ws.Cells[1, ($c + 1)]
            $cell.Value = $headers[$c]
        }
        Set-RangeStyle $ws "A1:K1" -BgColor $HeaderBg -FontColor $colorTekstWit -Bold $true -FontSize 11 -HAlign "Center"
        $ws.Row(1).Height = 22

        # Data-rijen
        $rowNum = 2
        foreach ($item in $Data) {
            $ws.Cells[$rowNum, 1].Value  = $item.Status
            $ws.Cells[$rowNum, 2].Value  = $item.DisplayName
            $ws.Cells[$rowNum, 3].Value  = $item.UserPrincipalName
            $ws.Cells[$rowNum, 4].Value  = $item.Mail
            $ws.Cells[$rowNum, 5].Value  = $item.EmployeeID
            $ws.Cells[$rowNum, 6].Value  = $item.Department
            $ws.Cells[$rowNum, 7].Value  = $item.JobTitle
            $ws.Cells[$rowNum, 8].Value  = if ($item.AccountEnabled) { "Ja" } else { "Nee" }
            $ws.Cells[$rowNum, 9].Value  = if ($item.HeeftKwalificerndeLic) { "✔  Ja" } else { "✖  Nee" }
            $ws.Cells[$rowNum, 10].Value = $item.KwalificerndeLicenties
            $ws.Cells[$rowNum, 11].Value = $item.AlleLicenties

            # Rijkleur op basis van licentie-status
            $rowBg = if ($item.HeeftKwalificerndeLic) { $StatusBgTrue } else { $StatusBgFalse }
            Set-RangeStyle $ws "A${rowNum}:K${rowNum}" -BgColor $rowBg -Border $true

            # Statuscel extra vet + kleur
            if ($item.HeeftKwalificerndeLic) {
                Set-RangeStyle $ws "A${rowNum}" -FontColor $colorGroen -Bold $true
                Set-RangeStyle $ws "I${rowNum}" -FontColor $colorGroen -Bold $true
            } else {
                Set-RangeStyle $ws "A${rowNum}" -FontColor $colorRood -Bold $true
                Set-RangeStyle $ws "I${rowNum}" -FontColor $colorRood -Bold $true
            }

            # Tekstomloop voor licentiekolommen
            $ws.Cells[$rowNum, 10].Style.WrapText = $true
            $ws.Cells[$rowNum, 11].Style.WrapText = $true
            $ws.Row($rowNum).Height = 30

            $rowNum++
        }

        # AutoFilter op headerrij
        $ws.Cells[1, 1, ($rowNum - 1), $headers.Count].AutoFilter = $true

        # Kolombreedtes
        Set-ColumnWidths $ws @{
            1 = 32   # Status
            2 = 28   # Naam
            3 = 36   # UPN
            4 = 30   # Mail
            5 = 14   # EmployeeID
            6 = 22   # Afdeling
            7 = 22   # Functie
            8 = 14   # Account actief
            9 = 20   # MDI licentie
            10= 42   # Kwalificerende lic
            11= 50   # Alle licenties
        }

        # Lettertype voor volledig werkblad
        $ws.Cells.Style.Font.Name = "Calibri"
        $ws.Cells.Style.Font.Size = 10

        return $ws
    }

    # ── Tabblad 2: Alle gebruikers ─────────────────────────────────────────
    Write-Step "  Tabblad 2: Alle gebruikers..."
    $sortedAll = $results | Sort-Object HeeftKwalificerndeLic, DisplayName
    Add-DataSheet -Pkg $excelPkg -SheetName "Alle gebruikers" -Data $sortedAll `
        -HeaderBg $colorBlauw `
        -AltRowBg $colorBlauwLicht `
        -StatusBgTrue  $colorGroenLicht `
        -StatusBgFalse $colorRoodLicht | Out-Null

    # ── Tabblad 3: Aanschaf vereist ────────────────────────────────────────
    Write-Step "  Tabblad 3: Aanschaf vereist..."
    Add-DataSheet -Pkg $excelPkg -SheetName "Aanschaf vereist" -Data $needsLicense `
        -HeaderBg $colorRood `
        -AltRowBg $colorRoodLicht `
        -StatusBgTrue  $colorGroenLicht `
        -StatusBgFalse $colorRoodLicht | Out-Null

    # ── Tabblad 4: Licentie aanwezig ──────────────────────────────────────
    Write-Step "  Tabblad 4: Licentie aanwezig..."
    Add-DataSheet -Pkg $excelPkg -SheetName "Licentie aanwezig" -Data $hasLicense `
        -HeaderBg $colorGroen `
        -AltRowBg $colorGroenLicht `
        -StatusBgTrue  $colorGroenLicht `
        -StatusBgFalse $colorRoodLicht | Out-Null

    # ── Tabvolgorde: Dashboard als eerste ─────────────────────────────────
    $excelPkg.Workbook.Worksheets.MoveToStart("Dashboard")

    # ── Opslaan ───────────────────────────────────────────────────────────
    $excelPkg.Save()
    $excelPkg.Dispose()

    Write-OK "Excel-rapport opgeslagen: $exportFilePath"

    # Stap 7: Afsluitrapport
    Write-Header "Conclusie"

    if ($needsLicense.Count -eq 0) {
        Write-OK "Alle in-scope gebruikers beschikken over een kwalificerende MDI-licentie."
        Write-OK "Er hoeven geen extra licenties aangeschaft te worden."
    } else {
        Write-Warn "$($needsLicense.Count) gebruiker(s) missen een kwalificerende licentie."
        Write-Warn "Raadpleeg het tabblad 'Aanschaf vereist' in het Excel-rapport:"
        Write-Host "  $exportFilePath" -ForegroundColor White
        Write-Host ""
        Write-Host "  Kwalificerende licenties voor MDI zijn:" -ForegroundColor Yellow
        Write-Host "    • Enterprise Mobility + Security E5/A5"
        Write-Host "    • Microsoft 365 A5 / E5 / G5"
        Write-Host "    • Microsoft 365 A5 / E5 / G5 / F5 Security"
        Write-Host "    • Microsoft 365 Business Premium"
        Write-Host "    • Microsoft 365 F5 Security + Compliance"
        Write-Host "    • Microsoft Defender for Office 365 Plan 1 (standalone)"
        Write-Host "    • Microsoft Defender Suite for Microsoft 365 Business Premium"
        Write-Host "    • + elke andere licentie die het Service Plan 'Microsoft Defender"
        Write-Host "      for Identity' bevat (wordt automatisch gedetecteerd)"
    }

    # Stap 8: Afmelden
    Disconnect-MgGraph | Out-Null
    Write-OK "Afgemeld van Microsoft Graph."
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
Main
