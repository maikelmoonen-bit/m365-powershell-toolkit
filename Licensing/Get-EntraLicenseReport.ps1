#Requires -Modules Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Genereert een licentie-rapport van alle actieve Entra ID licenties (zoals in de GUI).

.DESCRIPTION
    Haalt via Microsoft Graph alle actieve Entra ID licenties op (CapabilityStatus = Enabled),
    gebruikt de exacte GUI-namen zoals weergegeven in Azure > Entra ID > Licenses > All Products,
    en exporteert het resultaat als .docx tabel met alleen Licentienaam en Toegewezen.

.PARAMETER OutputPath
    Pad waar het .docx bestand wordt opgeslagen. Standaard: huidige map.

.PARAMETER TenantId
    Optioneel: Azure Tenant ID. Indien weggelaten wordt interactief ingelogd.

.EXAMPLE
    .\Get-EntraLicenseReport.ps1
    .\Get-EntraLicenseReport.ps1 -OutputPath "C:\Rapporten" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.NOTES
    Vereiste permissies: Organization.Read.All en Directory.Read.All
    Installeer modules via: Install-Module Microsoft.Graph -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [string]$TenantId
)

#region ─── SKU mapping: exacte GUI-namen (Azure > Entra ID > Licenses > All Products) ──
# Gebaseerd op de officiële Microsoft product names reference, aangevuld met
# de namen zoals de Entra-portal ze weergeeft.
$SkuFriendlyNames = @{
    # ── Microsoft 365 / Office 365 ──────────────────────────────────────────────
    "SPE_E3"                                       = "Microsoft 365 E3"
    "SPE_E5"                                       = "Microsoft 365 E5"
    "SPE_F1"                                       = "Microsoft 365 F3"
    "M365_F1"                                      = "Microsoft 365 F1"
    "SPB"                                          = "Microsoft 365 Business Premium"
    "O365_BUSINESS_ESSENTIALS"                     = "Microsoft 365 Business Basic"
    "SMB_BUSINESS_ESSENTIALS"                      = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"                        = "Microsoft 365 Business Standard"
    "SMB_BUSINESS_PREMIUM"                         = "Microsoft 365 Business Standard"
    "O365_BUSINESS"                                = "Microsoft 365 Apps for Business"
    "SMB_BUSINESS"                                 = "Microsoft 365 Apps for Business"
    "OFFICESUBSCRIPTION"                           = "Microsoft 365 Apps for Enterprise"
    "ENTERPRISEPACK"                               = "Office 365 E3"
    "ENTERPRISEPREMIUM"                            = "Office 365 E5"
    "ENTERPRISEPREMIUM_NOPSTNCONF"                 = "Office 365 E5 Without Audio Conferencing"
    "STANDARDPACK"                                 = "Office 365 E1"
    "DESKLESSPACK"                                 = "Office 365 F3"
    "STANDARDWOFFPACK"                             = "Office 365 E2"
    "ENTERPRISEWITHSCAL"                           = "Office 365 E4"
    "MIDSIZEPACK"                                  = "Office 365 Midsize Business"
    "LITEPACK"                                     = "Office 365 Small Business"
    "LITEPACK_P2"                                  = "Office 365 Small Business Premium"
    "OFFICE_PRO_PLUS_SUBSCRIPTION1"                = "Office 365 ProPlus"
    "STANDARDWOFFPACK_FACULTY"                     = "Office 365 A1 for faculty"
    "STANDARDWOFFPACK_STUDENT"                     = "Office 365 A1 for students"
    "STANDARDWOFFPACK_IW_FACULTY"                  = "Office 365 A3 for faculty"
    "STANDARDWOFFPACK_IW_STUDENT"                  = "Office 365 A3 for students"

    # ── Microsoft Entra ID (voorheen Azure AD) ───────────────────────────────────
    "AAD_BASIC"                                    = "Microsoft Entra ID Basic"
    "AAD_PREMIUM"                                  = "Azure Active Directory Premium P1"
    "AAD_PREMIUM_P2"                               = "Azure Active Directory Premium P2"
    "RIGHTSMANAGEMENT"                             = "Azure Information Protection Plan 1"

    # ── Exchange Online ──────────────────────────────────────────────────────────
    "EXCHANGE_S_STANDARD"                          = "Exchange Online (Plan 1)"
    "EXCHANGE_L_STANDARD"                          = "Exchange Online (Plan 1)"
    "EXCHANGESTANDARD"                             = "Exchange Online (Plan 1)"
    "EXCHANGE_S_ENTERPRISE"                        = "Exchange Online (Plan 2)"
    "EXCHANGE_ENTERPRISE"                          = "Exchange Online (Plan 2)"
    "EXCHANGEENTERPRISE"                           = "Exchange Online (Plan 2)"
    "EXCHANGE_S_DESKLESS"                          = "Exchange Online Kiosk"
    "EXCHANGE_S_ESSENTIALS"                        = "Exchange Online Essentials"
    "EXCHANGE_S_STANDARD_MIDMARKET"                = "Exchange Online (Plan 1)"

    # ── SharePoint / OneDrive ────────────────────────────────────────────────────
    "SHAREPOINTENTERPRISE"                         = "SharePoint Online (Plan 2)"
    "SHAREPOINTSTANDARD"                           = "SharePoint Online (Plan 1)"
    "WACONEDRIVEENTERPRISE"                        = "OneDrive for Business (Plan 2)"
    "WACONEDRIVESTANDARD"                          = "OneDrive for Business (Plan 1)"
    "ONEDRIVE_BASIC"                               = "OneDrive for Business (Plan 1)"
    "SHAREPOINTSTORAGE"                            = "SharePoint Online Storage"

    # ── Microsoft Teams ──────────────────────────────────────────────────────────
    "MCOCAP"                                       = "Microsoft Teams Shared Devices"
    "MCOTEAMS_ESSENTIALS"                          = "Microsoft Teams Essentials"
    "MS_TEAMS_IW"                                  = "Microsoft Teams Exploratory"
    "TEAMS_COMMERCIAL_TRIAL"                       = "Microsoft Teams Commercial Cloud Trial"
    "TEAMS_EXPLORATORY"                            = "Microsoft Teams Exploratory"
    "Microsoft_Teams_Premium"                      = "Microsoft Teams Premium"
    "MICROSOFT_TEAMS_ROOMS_PRO"                    = "Microsoft Teams Rooms Pro"
    "Teams_Premium_Dept"                           = "Teams Premium (For Departments)"
    "PHONESYSTEM_VIRTUALUSER"                      = "Microsoft Teams Phone Resource Account"
    "MCOPSTN1"                                     = "Microsoft 365 Domestic Calling Plan"
    "MCOPSTN2"                                     = "Microsoft 365 Domestic and International Calling Plan"
    "BUSINESS_VOICE_MED2_TELCO"                    = "Microsoft Teams Phone with Calling Plan"
    "MCOSTANDARD"                                  = "Skype for Business Online (Plan 2)"

    # ── Microsoft 365 Copilot ────────────────────────────────────────────────────
    "Microsoft_365_Copilot"                        = "Microsoft 365 Copilot"
    "Microsoft_Copilot_For_Finance"                = "Microsoft 365 Copilot for Finance (Preview)"

    # ── Dynamics 365 ────────────────────────────────────────────────────────────
    "DYN365_ENTERPRISE_CUSTOMER_SERVICE"           = "Dynamics 365 Customer Service Enterprise"
    "DYN365_ENTERPRISE_P1"                         = "Dynamics 365 P1 Trial for Information Workers"
    "DYN365_ENTERPRISE_P1_IW"                      = "Dynamics 365 P1 Trial for Information Workers"
    "DYN365_FINANCE"                               = "Dynamics 365 Finance"
    "DYN365_OPERATIONS_ACTIVITY"                   = "Dynamics 365 Operations - Activity"
    "DYN365_PROJECT_OPERATIONS_ATTACH"             = "Dynamics 365 Project Operations Attach"
    "DYN365_SCM_ATTACH"                            = "Dynamics 365 Supply Chain Management Attach to Qualifying Dynamics 365 Base Offer"
    "DYN365_TEAM_MEMBERS"                          = "Dynamics 365 Team Members"
    "DYN365_SALES_ENT_ATTACH"                      = "Dynamics 365 Sales Enterprise Attach to Qualifying Dynamics 365 Base Offer"
    "D365_SALES_ENT_ATTACH"                        = "Dynamics 365 Sales Enterprise Attach to Qualifying Dynamics 365 Base Offer"
    "DYN365_ENTERPRISE_SALES"                      = "Dynamics 365 Sales Enterprise Edition"
    "CRMSTANDARD"                                  = "Dynamics 365 Sales Enterprise Edition"
    "DYN365_CUSTOMER_INSIGHTS_VIRAL"               = "Dynamics 365 Customer Insights vTrial"
    "DYN365_CUSTOMER_INSIGHTS_ATTACH"              = "Dynamics 365 Customer Insights Attach New"
    "Dynamics_365_Sales_Premium_Viral_Trial"       = "Dynamics 365 Sales Premium Viral Trial"
    "DYN365_ENTERPRISE_PLAN1"                      = "Dynamics 365 Customer Engagement Plan"
    "DYN365_ENTERPRISE_TALENT"                     = "Dynamics 365 for Talent"
    "MSDYN365_SALES_INSIGHTS"                      = "Dynamics 365 Sales Insights"
    "FORMS_PRO"                                    = "Dynamics 365 Customer Voice Trial"
    "MICROSOFT_FORMS_PRO"                          = "Microsoft Dynamics 365 Customer Voice"
    "MICROSOFT_DYNAMICS_365_CUSTOMER_VOICE"        = "Microsoft Dynamics 365 Customer Voice"
    "AX7_USER_TRIAL"                               = "Microsoft Dynamics AX7 User Trial"
    "DYN365_FOR_OPERATIONS_SANDBOX_TIER2_SKU"      = "Dynamics 365 For Operations Sandbox Tier 2"
    "CCIBOTS_PRIVPREV_VIRAL"                       = "Microsoft Copilot Studio Viral Trial"

    # ── Power Platform ───────────────────────────────────────────────────────────
    "POWER_BI_STANDARD"                            = "Microsoft Fabric (Free)"
    "POWER_BI_PRO"                                 = "Power BI Pro"
    "PBI_PREMIUM_PER_USER"                         = "Power BI Premium Per User"
    "COLUMNAR_STORE"                               = "Power BI Premium Per User"
    "POWER_BI_ADDON"                               = "Power BI for Office 365 Add-On"
    "POWERAPPS_DEV"                                = "Microsoft Power Apps for Developer"
    "POWERAPPS_VIRAL"                              = "Microsoft Power Apps Plan 2 Trial"
    "POWERAPPS_PER_APP_IW"                         = "PowerApps per app baseline access"
    "POWERAPPS_PER_APP_NEW"                        = "Power Apps per app plan (new)"
    "POWERAPPS_PER_USER"                           = "Power Apps per user plan"
    "FLOW_FREE"                                    = "Microsoft Power Automate Free"
    "FLOW_PER_USER"                                = "Power Automate per user plan"
    "POWERAUTOMATE_ATTENDED_RPA"                   = "Power Automate Premium"
    "FLOW_P1"                                      = "Power Automate per user plan"
    "FLOW_P2"                                      = "Power Automate per user plan with attended RPA"
    "Power_Pages_vTrial_for_Makers"                = "Power Pages vTrial for Makers"
    "Power_Virtual_Agents"                         = "Power Virtual Agents"
    "VIRTUAL_AGENT_USL"                            = "Microsoft Copilot Studio User License"
    "CDSAICAPACITY"                                = "AI Builder Capacity add-on"
    "CDS_DB_CAPACITY"                              = "Common Data Service Database Capacity"
    "CDS_LOG_CAPACITY"                             = "Common Data Service Log Capacity"

    # ── Intune / Endpoint ────────────────────────────────────────────────────────
    "INTUNE_A"                                     = "Intune"
    "INTUNE_A_D"                                   = "Microsoft Intune Device"
    "INTUNE_A_VL"                                  = "Intune"
    "INTUNE_ENDPOINT_PRIVILEGE_MANAGEMENT"         = "Microsoft Intune Endpoint Privilege Management"
    "Microsoft_Intune_Endpoint_Privilege_Manageme" = "Microsoft Intune Endpoint Privilege Management"

    # ── Security / Compliance / Defender ────────────────────────────────────────
    "EMS"                                          = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"                                   = "Enterprise Mobility + Security E5"
    "INFORMATION_PROTECTION_COMPLIANCE"            = "Microsoft 365 E5 Compliance"
    "ADALLOM_STANDALONE"                           = "Microsoft Defender for Cloud Apps"
    "ATP_ENTERPRISE"                               = "Microsoft Defender for Office 365 (Plan 1)"
    "THREAT_INTELLIGENCE"                          = "Microsoft Defender for Office 365 (Plan 2)"
    "MDE_LITE"                                     = "Microsoft Defender for Endpoint P1"
    "MDE_SMB"                                      = "Microsoft Defender for Business"
    "MDATP_Server"                                 = "Microsoft Defender for Endpoint Server"
    "MFA_STANDALONE"                               = "Microsoft Azure Multi-Factor Authentication"

    # ── Project / Visio ──────────────────────────────────────────────────────────
    "PROJECTPROFESSIONAL"                          = "Planner and Project Plan 3"
    "PROJECTPREMIUM"                               = "Project Plan 5"
    "PROJECTONLINE_PLAN_1"                         = "Project Online Plan 1"
    "PROJECTONLINE_PLAN_2"                         = "Project Online Plan 2"
    "PROJECTCLIENT"                                = "Project Online Desktop Client"
    "PROJECTESSENTIALS"                            = "Project Online Essentials"
    "PROJECT_P1"                                   = "Planner Plan 1"
    "VISIOCLIENT"                                  = "Visio Plan 2"
    "VISIOONLINE_PLAN1"                            = "Visio Plan 1"

    # ── Overig ───────────────────────────────────────────────────────────────────
    "STREAM"                                       = "Microsoft Stream Trial"
    "Microsoft_Stream"                             = "Microsoft Stream Trial"
    "YAMMER_ENTERPRISE"                            = "Yammer Enterprise"
    "WINDOWS_STORE"                                = "Windows Store for Business"
    "WIN10_PRO_ENT_SUB"                            = "Windows 10 Enterprise E3"
    "WIN10_VDA_E3"                                 = "Windows 10 Enterprise E3"
    "WIN10_VDA_E5"                                 = "Windows 10 Enterprise E5"
    "MICROSOFT_BUSINESS_CENTER"                    = "Microsoft Business Center"
    "SPZA_IW"                                      = "App Connect"
    "PROJECT_MADEIRA_PREVIEW_IW_SKU"               = "Dynamics 365 Business Central for IWs"
    "SMB_APPS"                                     = "Business Apps (free)"
    "Microsoft_Cloud_For_Sustainability_Vtrial"    = "Microsoft Cloud for Sustainability vTrial"
    "Microsoft_Cloud_PKI"                          = "Microsoft Cloud PKI"
    "ECAL_SERVICES"                                = "ECAL Services"
    "EXCHANGE_ANALYTICS"                           = "Microsoft MyAnalytics"
    "WORKPLACE_ANALYTICS"                          = "Microsoft Workplace Analytics"
    "MICROSOFT_REMOTE_ASSIST"                      = "Dynamics 365 Remote Assist"
    "IT_ACADEMY_AD"                                = "Microsoft Imagine Academy"
    "M365_G3_GOV"                                  = "Microsoft 365 G3 GCC"
}
#endregion

#region ─── Functies ─────────────────────────────────────────────────────────────
function Get-FriendlyLicenseName {
    param([string]$SkuPartNumber)
    if ($SkuFriendlyNames.ContainsKey($SkuPartNumber)) {
        return $SkuFriendlyNames[$SkuPartNumber]
    }
    # Fallback: maak de SKU-naam leesbaar
    $readable = $SkuPartNumber -replace "_", " "
    return (Get-Culture).TextInfo.ToTitleCase($readable.ToLower())
}

function Write-Status {
    param([string]$Message, [string]$Color = "Cyan")
    Write-Host "  $Message" -ForegroundColor $Color
}
#endregion

#region ─── Verbinding maken met Microsoft Graph ─────────────────────────────────
Clear-Host
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Blue
Write-Host "║        Entra ID Licentie Rapport Generator                  ║" -ForegroundColor Blue
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Blue
Write-Host ""

if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Identity.DirectoryManagement")) {
    Write-Host "  [!] Microsoft.Graph module niet gevonden." -ForegroundColor Yellow
    Write-Host "  Installeer via: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
    $install = Read-Host "  Nu installeren? (j/n)"
    if ($install -eq "j") { Install-Module Microsoft.Graph -Scope CurrentUser -Force }
    else { exit 1 }
}

Write-Status "Verbinding maken met Microsoft Graph..." "Cyan"
$connectParams = @{ Scopes = @("Organization.Read.All", "Directory.Read.All"); NoWelcome = $true }
if ($TenantId) { $connectParams.TenantId = $TenantId }

try {
    Connect-MgGraph @connectParams -ErrorAction Stop
    Write-Status "Verbonden!" "Green"
} catch {
    Write-Host "  [FOUT] Verbinding mislukt: $_" -ForegroundColor Red
    exit 1
}
#endregion

#region ─── Data ophalen ─────────────────────────────────────────────────────────
Write-Status "Licentiedata ophalen..." "Cyan"

try {
    $subscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    $org = Get-MgOrganization -ErrorAction Stop
    $tenantName = $org[0].DisplayName
    Write-Status "Organisatie: $tenantName" "Green"
} catch {
    Write-Host "  [FOUT] Data ophalen mislukt: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit 1
}

# Filter: ALLEEN actieve licenties (zelfde als de GUI — Enabled = zichtbaar, de rest niet)
$activeLicenses = $subscribedSkus | Where-Object { $_.CapabilityStatus -eq "Enabled" }
Write-Status "$($activeLicenses.Count) actieve licentie(s) gevonden" "Green"

# Verwerk naar overzicht met alleen naam + toegewezen, gesorteerd op naam
$licenseData = @(foreach ($sku in $activeLicenses) {
    [PSCustomObject]@{
        Licentie   = Get-FriendlyLicenseName -SkuPartNumber $sku.SkuPartNumber
        SKU        = $sku.SkuPartNumber
        Toegewezen = $sku.ConsumedUnits
    }
}) | Sort-Object Licentie

# Console preview
Write-Host ""
Write-Host ("  {0,-55} {1,12}" -f "Licentie", "Toegewezen") -ForegroundColor White
Write-Host ("  " + "-" * 70) -ForegroundColor DarkGray
foreach ($lic in $licenseData) {
    $displayName = if ($lic.Licentie.Length -gt 54) { $lic.Licentie.Substring(0,51) + "..." } else { $lic.Licentie }
    Write-Host ("  {0,-55} {1,12}" -f $displayName, $lic.Toegewezen) -ForegroundColor Gray
}
Write-Host ("  " + "-" * 70) -ForegroundColor DarkGray
$totalAssigned = ($licenseData | Measure-Object -Property Toegewezen -Sum).Sum
Write-Host ("  {0,-55} {1,12}" -f "Totaal", $totalAssigned) -ForegroundColor White
Write-Host ""
#endregion

#region ─── Word document aanmaken ───────────────────────────────────────────────
Write-Status "Word-document aanmaken..." "Cyan"

$npmCheck = Get-Command npm -ErrorAction SilentlyContinue
if (-not $npmCheck) {
    Write-Host "  [!] npm niet gevonden. Word-export overgeslagen." -ForegroundColor Yellow
    Write-Host "      Installeer Node.js via https://nodejs.org om .docx export te activeren." -ForegroundColor Yellow
} else {
    $tmpDir          = Join-Path $env:TEMP "EntraLicenseReport_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -ItemType Directory -Path $tmpDir | Out-Null

    $reportDate      = Get-Date -Format "MMMM yyyy"
    $reportDateShort = Get-Date -Format "yyyy-MM"
    $outputFile      = Join-Path $OutputPath "EntraID_Licenties_$reportDateShort.docx"

    # Alleen naam + toegewezen doorgeven aan Node.js
    $jsonData = $licenseData | Select-Object Licentie, Toegewezen | ConvertTo-Json -Compress

    $nodeScript = @'
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign } = require('docx');
const fs = require('fs');

const dataFile   = process.argv[2];
const outputFile = process.argv[3];
const tenantName = process.argv[4];
const reportDate = process.argv[5];

const licenses = JSON.parse(fs.readFileSync(dataFile, 'utf8'));

// Kleuren
const colorPrimary  = "1F3864";
const colorSecondary= "2E75B6";
const colorHeaderBg = "1F3864";
const colorRowAlt   = "EBF3FB";
const colorRowNorm  = "FFFFFF";
const colorBorder   = "BFBFBF";

// A4 staand, 1 cm marges: content width = 11906 - 2*567 = 10772 DXA
const COL_NAME   = 8772;
const COL_ASSIGN = 2000;
const TABLE_W    = COL_NAME + COL_ASSIGN;

function cell(text, opts = {}) {
    const { bold = false, bgColor = colorRowNorm, align = AlignmentType.LEFT,
            width = COL_NAME, isHeader = false, color = "000000" } = opts;
    const b = { style: BorderStyle.SINGLE, size: 1, color: colorBorder };
    return new TableCell({
        width: { size: width, type: WidthType.DXA },
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        verticalAlign: VerticalAlign.CENTER,
        borders: { top: b, bottom: b, left: b, right: b },
        children: [new Paragraph({
            alignment: align,
            children: [new TextRun({
                text: String(text),
                bold: bold || isHeader,
                color: isHeader ? "FFFFFF" : color,
                size: isHeader ? 20 : 19,
                font: "Calibri"
            })]
        })]
    });
}

const headerRow = new TableRow({
    tableHeader: true,
    children: [
        cell("Licentie",    { isHeader: true, bgColor: colorHeaderBg, width: COL_NAME }),
        cell("Toegewezen",  { isHeader: true, bgColor: colorHeaderBg, width: COL_ASSIGN, align: AlignmentType.CENTER }),
    ]
});

const dataRows = licenses.map((lic, i) => {
    const bg = i % 2 === 0 ? colorRowNorm : colorRowAlt;
    return new TableRow({
        children: [
            cell(lic.Licentie,                           { bgColor: bg, width: COL_NAME }),
            cell(lic.Toegewezen.toLocaleString('nl-NL'), { bgColor: bg, width: COL_ASSIGN, align: AlignmentType.CENTER }),
        ]
    });
});

const totalAssigned = licenses.reduce((s, l) => s + l.Toegewezen, 0);
const totalRow = new TableRow({
    children: [
        cell("Totaal",                                   { bold: true, bgColor: "D6E4F0", width: COL_NAME }),
        cell(totalAssigned.toLocaleString('nl-NL'),      { bold: true, bgColor: "D6E4F0", width: COL_ASSIGN, align: AlignmentType.CENTER }),
    ]
});

const border6 = { style: BorderStyle.SINGLE, size: 6, color: colorSecondary, space: 1 };

const doc = new Document({
    sections: [{
        properties: {
            page: {
                size: { width: 11906, height: 16838 },  // A4 staand
                margin: { top: 567, right: 567, bottom: 567, left: 567 }  // 1 cm
            }
        },
        children: [
            new Paragraph({
                spacing: { before: 0, after: 100 },
                border: { bottom: border6 },
                children: [
                    new TextRun({ text: "Licentie-overzicht  |  ", bold: true, size: 28, color: colorPrimary, font: "Calibri" }),
                    new TextRun({ text: tenantName, bold: false, size: 24, color: colorSecondary, font: "Calibri" }),
                    new TextRun({ text: "   -   " + reportDate, bold: false, size: 20, color: "888888", font: "Calibri" }),
                ]
            }),
            new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "", font: "Calibri" })] }),
            new Paragraph({
                spacing: { before: 0, after: 120 },
                children: [new TextRun({
                    text: `${licenses.length} actieve licentiesoorten  |  ${totalAssigned.toLocaleString('nl-NL')} licenties toegewezen`,
                    size: 18, color: "555555", font: "Calibri", italics: true
                })]
            }),
            new Table({
                width: { size: TABLE_W, type: WidthType.DXA },
                columnWidths: [COL_NAME, COL_ASSIGN],
                rows: [headerRow, ...dataRows, totalRow]
            }),
            new Paragraph({ spacing: { before: 160, after: 0 }, children: [new TextRun({ text: "", font: "Calibri" })] }),
            new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [new TextRun({
                    text: "Bron: Azure Entra ID > Licenses > All Products (alleen actieve licenties, CapabilityStatus = Enabled)",
                    size: 15, color: "999999", font: "Calibri", italics: true
                })]
            }),
        ]
    }]
});

Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(outputFile, buf);
    console.log('OK:' + outputFile);
}).catch(err => {
    console.error('ERR:' + err.message);
    process.exit(1);
});
'@

    $nodeScriptPath = Join-Path $tmpDir "make_report.js"
    $jsonDataPath   = Join-Path $tmpDir "data.json"

    Set-Content -Path $nodeScriptPath -Value $nodeScript -Encoding UTF8
    Set-Content -Path $jsonDataPath   -Value $jsonData   -Encoding UTF8

    Write-Status "npm pakket installeren..." "DarkGray"
    Push-Location $tmpDir
    & npm install docx --save-quiet 2>&1 | Out-Null

    $result = & node $nodeScriptPath $jsonDataPath $outputFile $tenantName $reportDate 2>&1
    Pop-Location

    if ($result -match "^OK:") {
        Write-Status "Word-document aangemaakt!" "Green"
        Write-Host ""
        Write-Host "  OK  Rapport opgeslagen:" -ForegroundColor Green
        Write-Host "      $outputFile" -ForegroundColor White
    } else {
        Write-Host "  [!] Fout bij aanmaken Word-document:" -ForegroundColor Yellow
        Write-Host "      $result" -ForegroundColor Yellow
    }

    Remove-Item -Path $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}
#endregion

#region ─── CSV export ────────────────────────────────────────────────────────────
$csvFile = Join-Path $OutputPath "EntraID_Licenties_$(Get-Date -Format 'yyyy-MM').csv"
$licenseData | Select-Object Licentie, Toegewezen |
    Export-Csv -Path $csvFile -Delimiter ";" -Encoding UTF8 -NoTypeInformation

Write-Host "  OK  CSV-backup opgeslagen:" -ForegroundColor Green
Write-Host "      $csvFile" -ForegroundColor White
#endregion

Disconnect-MgGraph | Out-Null
Write-Host ""
Write-Host "  Klaar." -ForegroundColor Green
Write-Host ""
