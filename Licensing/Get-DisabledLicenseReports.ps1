#Requires -Version 5.1
<#
.SYNOPSIS
    Entra ID Disabled Account License Audit
    Exporteert een opgemaakt Excel-rapport met uitgeschakelde accounts
    die nog licenties hebben toegewezen, inclusief prijsreferenties.

.DESCRIPTION
    Dit script:
      1. Verbindt met Microsoft Graph
      2. Haalt uitgeschakelde accounts op met toegewezen licenties
      3. Zet SKU GUIDs om naar SKU codes
      4. Genereert een Excel-rapport met prijsreferenties
      5. Markeert onbekende of legacy prijzen als "Onbekend" i.p.v. 0,00

.NOTES
    Vereisten:
      - PowerShell 5.1 of hoger
      - Microsoft.Graph module
      - Python 3.8+ met openpyxl
      - Graph scopes: User.Read.All, Directory.Read.All, Organization.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "",

    [Parameter(Mandatory = $false)]
    [string]$TenantId = "",

    [Parameter(Mandatory = $false)]
    [switch]$SkipPythonCheck
)

if (-not $OutputPath) {
    $OutputPath = ".\EntraID_Disabled_License_Report_$(Get-Date -Format 'yyyy-MM-dd_HHmm').xlsx"
}

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

function Find-Python {
    foreach ($cmd in @("python3", "python", "py")) {
        try {
            $v = & $cmd --version 2>&1
            if ("$v" -match "Python 3") { return $cmd }
        } catch {}
    }
    return $null
}

Clear-Host
Write-Host ""
Write-Host "  =================================================================" -ForegroundColor Cyan
Write-Host "    Entra ID Disabled Account License Audit" -ForegroundColor Cyan
Write-Host "    Licentiekosten van uitgeschakelde accounts" -ForegroundColor Cyan
Write-Host "  =================================================================" -ForegroundColor Cyan
Write-Host ""

# STAP 1 - PYTHON CONTROLEREN
$pythonCmd = $null
if (-not $SkipPythonCheck) {
    Write-Host "  [1/5] Python omgeving controleren..." -ForegroundColor White
    $pythonCmd = Find-Python
    if (-not $pythonCmd) {
        Write-Host ""
        Write-Host "  FOUT: Python 3 niet gevonden." -ForegroundColor Red
        Write-Host "  Download via: https://www.python.org/downloads/" -ForegroundColor Yellow
        exit 1
    }
    Write-Host "     OK: Python gevonden ($pythonCmd)" -ForegroundColor Green

    $chk = & $pythonCmd -c "import openpyxl; print('OK')" 2>&1
    if ("$chk" -ne "OK") {
        Write-Host "     openpyxl installeren..." -ForegroundColor Yellow
        & $pythonCmd -m pip install openpyxl --quiet 2>&1 | Out-Null
        & $pythonCmd -m pip install openpyxl --quiet --break-system-packages 2>&1 | Out-Null
    }

    $chk2 = & $pythonCmd -c "import openpyxl; print('OK')" 2>&1
    if ("$chk2" -ne "OK") {
        Write-Host "  FOUT: openpyxl kon niet worden geinstalleerd." -ForegroundColor Red
        exit 1
    }
    Write-Host "     OK: openpyxl beschikbaar" -ForegroundColor Green
}
else {
    $pythonCmd = Find-Python
    if (-not $pythonCmd) { $pythonCmd = "python" }
}

# STAP 2 - MICROSOFT.GRAPH CONTROLEREN
Write-Host ""
Write-Host "  [2/5] Microsoft Graph module controleren..." -ForegroundColor White

$graphMod = Get-Module -ListAvailable -Name "Microsoft.Graph.Users" -ErrorAction SilentlyContinue
if (-not $graphMod) {
    Write-Host "     Microsoft.Graph installeren..." -ForegroundColor Yellow
    try {
        Install-Module -Name "Microsoft.Graph" -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "     OK: Microsoft.Graph geinstalleerd" -ForegroundColor Green
    } catch {
        Write-Host "  FOUT: $_" -ForegroundColor Red
        exit 1
    }
}
else {
    $modVer = ($graphMod | Select-Object -First 1).Version
    Write-Host "     OK: Microsoft.Graph aanwezig (v$modVer)" -ForegroundColor Green
}

Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue

# STAP 3 - VERBINDEN MET GRAPH
Write-Host ""
Write-Host "  [3/5] Verbinden met Microsoft Graph..." -ForegroundColor White

$scopes = @("User.Read.All", "Directory.Read.All", "Organization.Read.All")

$tenantName = "Onbekende Tenant"
$tenantDomain = "N/A"

try {
    if ($TenantId) {
        Connect-MgGraph -Scopes $scopes -TenantId $TenantId -NoWelcome -ErrorAction Stop
    } else {
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
    }
} catch {
    Write-Host "  FOUT bij verbinden: $_" -ForegroundColor Red
    exit 1
}

try {
    $orgList = @(Get-MgOrganization -ErrorAction Stop)
    if ($orgList.Count -gt 0) {
        $org = $orgList[0]
        $tenantName = if ($org.DisplayName) { $org.DisplayName } else { "Onbekende Tenant" }

        $verDomains = @($org.VerifiedDomains)
        if ($verDomains.Count -gt 0) {
            $primary = $verDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -First 1
            if ($primary) {
                $tenantDomain = $primary.Name
            } else {
                $tenantDomain = $verDomains[0].Name
            }
        }
    }
    Write-Host "     OK: Verbonden met $tenantName ($tenantDomain)" -ForegroundColor Green
} catch {
    Write-Host "     Waarschuwing: tenantinfo kon niet worden opgehaald ($_)" -ForegroundColor Yellow
}

# STAP 4 - DATA OPHALEN
Write-Host ""
Write-Host "  [4/5] Uitgeschakelde accounts met licenties ophalen..." -ForegroundColor White

$allUsers = [System.Collections.Generic.List[object]]::new()

# SKU map ophalen
$skuMap = @{}
try {
    Write-Host "     SKU-namen ophalen..." -ForegroundColor Gray
    $skus = @(Get-MgSubscribedSku -ErrorAction Stop)
    foreach ($s in $skus) {
        $skuMap[$s.SkuId.ToString()] = $s.SkuPartNumber
    }
    Write-Host "     OK: $($skuMap.Count) SKU's gevonden in tenant" -ForegroundColor Green
} catch {
    Write-Host "     Waarschuwing: SKU-namen konden niet worden opgehaald ($_)" -ForegroundColor Yellow
}

# Gebruikers ophalen
try {
    $gFilter = "accountEnabled eq false and assignedLicenses/`$count ne 0"
    $gSelect = "id,displayName,userPrincipalName,assignedLicenses,department,jobTitle,mail,createdDateTime"

    $rawUsers = @(Get-MgUser `
        -Filter $gFilter `
        -Select $gSelect `
        -CountVariable gCount `
        -ConsistencyLevel eventual `
        -All `
        -PageSize 500 `
        -ErrorAction Stop)

    $total = $rawUsers.Count
    Write-Host "     $total uitgeschakelde accounts met licentie(s) gevonden" -ForegroundColor Cyan

    $idx = 0
    foreach ($u in $rawUsers) {
        $idx++
        if ($idx % 25 -eq 0 -or $idx -eq $total) {
            $pct = [int](($idx / [Math]::Max($total, 1)) * 100)
            Write-Progress -Activity "Gebruikers verwerken" -Status "$idx / $total" -PercentComplete $pct
        }

        $licNames = @()
        $assignedLics = @($u.AssignedLicenses)
        foreach ($lic in $assignedLics) {
            $guid = $lic.SkuId.ToString()
            if ($skuMap.ContainsKey($guid)) {
                $licNames += $skuMap[$guid]
            } else {
                $licNames += $guid
            }
        }

        $record = [PSCustomObject]@{
            displayName = if ($u.DisplayName)      { "$($u.DisplayName)" } else { "" }
            upn         = if ($u.UserPrincipalName) { "$($u.UserPrincipalName)" } else { "" }
            mail        = if ($u.Mail)             { "$($u.Mail)" } else { "" }
            department  = if ($u.Department)       { "$($u.Department)" } else { "" }
            jobTitle    = if ($u.JobTitle)         { "$($u.JobTitle)" } else { "" }
            created     = if ($u.CreatedDateTime)  { $u.CreatedDateTime.ToString("dd-MM-yyyy") } else { "" }
            licenses    = $licNames
        }

        $allUsers.Add($record)
    }

    Write-Progress -Completed -Activity "Klaar"
} catch {
    Write-Host "  FOUT bij ophalen gebruikers: $_" -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    exit 1
}

Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Host "     Graph verbinding gesloten" -ForegroundColor Green

# STAP 5 - EXCEL GENEREREN
Write-Host ""
Write-Host "  [5/5] Excel-rapport genereren..." -ForegroundColor White

$tempJson = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "entraid_audit_$(Get-Random).json")
$exportObj = [PSCustomObject]@{
    tenant   = $tenantName
    run_date = (Get-Date -Format "dd-MM-yyyy HH:mm")
    users    = $allUsers
}
$exportObj | ConvertTo-Json -Depth 10 | Set-Content -Path $tempJson -Encoding UTF8

$tempPy = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "entraid_excel_$(Get-Random).py")

$outputFull = [System.IO.Path]::GetFullPath($OutputPath)
$outputDir  = [System.IO.Path]::GetDirectoryName($outputFull)
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$pyCode = @'
import json
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

# ============================================================
# Actuele publieke NL-prijsreferenties (EUR per user/month)
# Unknown / legacy / niet publiek geprijsde SKU's blijven None
# ============================================================

PRICE_TABLE = {
    # Business
    "MICROSOFT_365_BUSINESS_BASIC":     {"name": "Microsoft 365 Business Basic",           "price": 5.20},
    "MICROSOFT_365_BUSINESS_STANDARD":  {"name": "Microsoft 365 Business Standard",        "price": 10.80},
    "MICROSOFT_365_BUSINESS_PREMIUM":   {"name": "Microsoft 365 Business Premium",         "price": 19.10},
    "M365_APPS_BUSINESS":               {"name": "Microsoft 365-apps voor bedrijven",      "price": 9.10},

    # Enterprise suites
    "ENTERPRISEPACK":                   {"name": "Office 365 E3",                          "price": 23.20},
    "ENTERPRISEPREMIUM":                {"name": "Office 365 E5",                          "price": 38.40},
    "SPE_E3":                           {"name": "Microsoft 365 E3",                       "price": 34.90},
    "SPE_E5":                           {"name": "Microsoft 365 E5",                       "price": 55.20},
    "M365_APPS_ENTERPRISE":             {"name": "Microsoft 365-apps voor ondernemingen", "price": 13.20},
    "STANDARDPACK":                     {"name": "Office 365 E1",                          "price": 8.70},

    # Frontline
    "M365_F1":                          {"name": "Microsoft 365 F1",                       "price": 1.90},
    "M365_F3":                          {"name": "Microsoft 365 F3",                       "price": 6.90},
    "O365_F3":                          {"name": "Office 365 F3",                          "price": 6.90},

    # Exchange
    "EXCHANGESTANDARD":                 {"name": "Exchange Online Plan 1",                 "price": 3.50},
    "EXCHANGEENTERPRISE":               {"name": "Exchange Online Plan 2",                 "price": 6.90},

    # Identity / security / management
    "AAD_PREMIUM":                      {"name": "Microsoft Entra ID P1",                  "price": 5.20},
    "AAD_PREMIUM_P2":                   {"name": "Microsoft Entra ID P2",                  "price": 7.80},
    "INTUNE_A":                         {"name": "Microsoft Intune",                       "price": 6.90},
    "EMS":                              {"name": "Enterprise Mobility + Security E3",      "price": 9.20},
    "EMSPREMIUM":                       {"name": "Enterprise Mobility + Security E5",      "price": 14.20},

    # Teams / voice
    "MCOSTANDARD":                      {"name": "Teams Phone Standard",                   "price": 8.70},
    "PHONESYSTEM_VIRTUALUSER":          {"name": "Teams Phone Resource Account",           "price": 0.00},

    # Power Platform / BI
    "POWER_BI_STANDARD":                {"name": "Power BI Free",                          "price": 0.00},
    "POWER_BI_PRO":                     {"name": "Power BI Pro",                           "price": 12.10},
    "POWER_BI_PREMIUM_PER_USER":        {"name": "Power BI Premium Per User",              "price": 20.80},
    "FLOW_FREE":                        {"name": "Power Automate Free",                    "price": 0.00},
    "POWERFLOW_P1":                     {"name": "Power Automate Premium",                 "price": 13.00},
    "POWERAPPS_PREMIUM":                {"name": "Power Apps Premium",                     "price": 17.30},

    # Planner / Project
    "PROJECT_P1":                       {"name": "Planner Abonnement 1",                   "price": 8.70},
    "PROJECT_P3":                       {"name": "Planner en Project Abonnement 3",        "price": 26.00},
    "PROJECT_P5":                       {"name": "Planner en Project Abonnement 5",        "price": 47.70},

    # Visio
    "VISIO_PLAN1":                      {"name": "Visio Abonnement 1",                     "price": 4.30},
    "VISIOCLIENT":                      {"name": "Visio Abonnement 2",                     "price": 13.00},
}

# Aliases en oude SKU-benamingen -> canonieke prijs-SKU
SKU_ALIASES = {
    # Oude naar nieuwe business naming
    "O365_BUSINESS_ESSENTIALS": "MICROSOFT_365_BUSINESS_BASIC",
    "O365_BUSINESS_PREMIUM":    "MICROSOFT_365_BUSINESS_STANDARD",
    "SMB_BUSINESS":             "M365_APPS_BUSINESS",

    # Enterprise
    "ENTERPRISEPACK":          "ENTERPRISEPACK",
    "ENTERPRISEPREMIUM":       "ENTERPRISEPREMIUM",
    "SPE_E3":                  "SPE_E3",
    "SPE_E5":                  "SPE_E5",
    "STANDARDPACK":            "STANDARDPACK",

    # Microsoft 365 apps
    "OFFICESUBSCRIPTION":      "M365_APPS_ENTERPRISE",
    "MCOMEETADV_FACULTY":      "M365_APPS_ENTERPRISE",

    # Frontline / deskless
    "O365_ENTERPRISEK1":       "M365_F1",
    "DESKLESSPACK":            "M365_F1",
    "O365_ENTERPRISEK2":       "O365_F3",
    "SPE_F1":                  "M365_F1",
    "SPE_F3":                  "M365_F3",

    # Exchange
    "EXCHANGESTANDARD":        "EXCHANGESTANDARD",
    "EXCHANGEENTERPRISE":      "EXCHANGEENTERPRISE",

    # Identity / security / management
    "AAD_PREMIUM":             "AAD_PREMIUM",
    "AAD_PREMIUM_P2":          "AAD_PREMIUM_P2",
    "INTUNE_A":                "INTUNE_A",
    "EMS":                     "EMS",
    "EMSPREMIUM":              "EMSPREMIUM",

    # Teams / voice
    "MCOSTANDARD":             "MCOSTANDARD",
    "PHONESYSTEM_VIRTUALUSER": "PHONESYSTEM_VIRTUALUSER",

    # BI / Power Platform
    "POWER_BI_STANDARD":             "POWER_BI_STANDARD",
    "POWER_BI_PRO":                  "POWER_BI_PRO",
    "POWER_BI_PREMIUM_PER_USER":     "POWER_BI_PREMIUM_PER_USER",
    "FLOW_FREE":                     "FLOW_FREE",
    "POWERFLOW_P1":                  "POWERFLOW_P1",
    "POWERAPPS_PER_USER":            "POWERAPPS_PREMIUM",

    # Legacy Power Apps per app - bewust onbekend houden
    "POWERAPPS_PER_APP_IW":          "POWERAPPS_PER_APP_IW",
    "POWERAPPS_PER_APP":             "POWERAPPS_PER_APP",

    # Legacy project mappings
    "PROJECT_P1":               "PROJECT_P1",
    "PROJECT_P2":               "PROJECT_P3",
    "PROJECT_P3":               "PROJECT_P5",
    "PROJECTPROFESSIONAL":      "PROJECT_P3",

    # Visio
    "VISIO_PLAN1":              "VISIO_PLAN1",
    "VISIOCLIENT":              "VISIOCLIENT",

    # Other / passthrough
    "DYN365_OPERATIONS_ACTIVITY": "DYN365_OPERATIONS_ACTIVITY",
    "FORMS_PRO":                 "FORMS_PRO",
}

# Fallback friendly names als er geen prijsentry bestaat
FRIENDLY_NAMES = {
    "POWERAPPS_PER_APP_IW":      "Power Apps per app (legacy)",
    "POWERAPPS_PER_APP":         "Power Apps per app",
    "DYN365_OPERATIONS_ACTIVITY":"Dynamics 365 Operations - Activity",
    "FORMS_PRO":                 "Forms Pro",
}

C = {
    "db":"0D2137","mb":"0F3460","ac":"E94560","or":"F5A623","gr":"27AE60",
    "lg":"F2F4F7","mg":"BDC3C7","wh":"FFFFFF","dm":"7F8C8D","dg":"C0392B","ar":"EBF5FB",
    "yl":"FFF3CD"
}

def canonical_sku(sk):
    sk = (sk or "").upper().strip()
    return SKU_ALIASES.get(sk, sk)

def price_for_sku(sk):
    csk = canonical_sku(sk)
    info = PRICE_TABLE.get(csk)
    return None if info is None else info.get("price")

def name_for_sku(sk):
    csk = canonical_sku(sk)
    info = PRICE_TABLE.get(csk)
    if info and info.get("name"):
        return info["name"]
    return FRIENDLY_NAMES.get(csk, sk)

def has_known_price(sk):
    return price_for_sku(sk) is not None

def f(h): return PatternFill("solid", fgColor=h)
def ft(b=False, s=11, c="1A1A2E", i=False): return Font(bold=b, size=s, color=c, name="Arial", italic=i)
def ac(w=False): return Alignment(horizontal="center", vertical="center", wrap_text=w)
def al(w=False): return Alignment(horizontal="left", vertical="center", wrap_text=w)
def bd():
    t = Side(style="thin", color="BDC3C7")
    return Border(left=t, right=t, top=t, bottom=t)
def sw(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w

def ms(ws, rng, txt, fill, fon, aln, rh=None):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = txt
    c.fill = f(fill)
    c.font = fon
    c.alignment = aln
    if rh:
        rn = int("".join(x for x in rng.split(":")[0] if x.isdigit()))
        ws.row_dimensions[rn].height = rh

def set_cell(cell, value, fill_color, align_left=False, bold=False, font_color="1A1A2E", size=10):
    cell.value = value
    cell.fill = f(fill_color)
    cell.font = ft(b=bold, s=size, c=font_color)
    cell.alignment = al(True) if align_left else ac(True)
    cell.border = bd()

def set_currency_or_text(cell, value, fill_color, align_left=False, bold=False, font_color="1A1A2E", size=10):
    set_cell(cell, value, fill_color, align_left=align_left, bold=bold, font_color=font_color, size=size)
    if isinstance(value, (int, float)):
        cell.number_format = '#,##0.00'

def build_summary(users):
    lc = {}
    unknown_skus = {}
    for u in users:
        for raw_sk in (u.get("licenses") or []):
            sk = (raw_sk or "").upper().strip()
            lc[sk] = lc.get(sk, 0) + 1
            if price_for_sku(sk) is None:
                unknown_skus[sk] = unknown_skus.get(sk, 0) + 1

    known_monthly_by_sku = {}
    for sk, cnt in lc.items():
        pr = price_for_sku(sk)
        if pr is not None:
            known_monthly_by_sku[sk] = pr * cnt

    known_month_total = sum(known_monthly_by_sku.values())
    known_year_total = known_month_total * 12
    unknown_assignment_count = sum(unknown_skus.values())

    return lc, unknown_skus, known_monthly_by_sku, known_month_total, known_year_total, unknown_assignment_count

def dashboard(ws, tenant, run_date, users, lc, known_monthly_by_sku, known_month_total, known_year_total, unknown_assignment_count):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90

    for i, w in enumerate([2,18,18,18,18,18,18,18,2], 1):
        sw(ws, i, w)

    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 58
    ms(ws, "B2:H2", "Microsoft Entra ID  |  Disabled Account License Audit",
       C["db"], ft(True,20,C["wh"]), ac(), 58)

    ws.row_dimensions[3].height = 20
    ms(ws, "B3:H3", f"Tenant: {tenant}   |   Rapport: {run_date}   |   Prijsbron: publieke Microsoft NL referentieprijzen",
       C["mb"], ft(False,10,C["mg"]), ac())

    ws.row_dimensions[4].height = 5
    ms(ws, "B4:H4", "", C["ac"], ft(), ac())

    tu = len(users)
    tl = sum(lc.values())

    ws.row_dimensions[6].height = 50
    ws.row_dimensions[7].height = 20

    cards = [
        ("B6:C6", str(tu), C["ac"], C["wh"], "B7:C7", "Uitgeschakelde Accounts"),
        ("D6:E6", str(tl), C["mb"], C["wh"], "D7:E7", "Gekoppelde Licenties"),
        ("F6:G6", f"EUR {known_month_total:,.2f}", C["or"], "1A1A2E", "F7:G7", "Bekende Mnd. Kosten"),
        ("H6:H6", f"EUR {known_year_total:,.2f}", C["dg"], C["wh"], "H7:H7", "Bekende Jaar Kosten")
    ]
    for top_rng, val, bg, fg, lbl_rng, lbl in cards:
        ws.merge_cells(top_rng)
        c = ws[top_rng.split(":")[0]]
        c.value = val
        c.fill = f(bg)
        c.font = ft(True,26,fg)
        c.alignment = ac()

        ws.merge_cells(lbl_rng)
        c2 = ws[lbl_rng.split(":")[0]]
        c2.value = lbl
        c2.fill = f(bg)
        c2.font = ft(False,9,fg)
        c2.alignment = ac(True)

    ws.row_dimensions[8].height = 8
    ms(ws, "B9:H9", "  Top Licenties bij Uitgeschakelde Gebruikers",
       C["mb"], ft(True,12,C["wh"]), al(), 26)

    headers = ["Licentienaam","SKU Code","# Gebruikers","EUR p.u./mnd","Maand. EUR","Jaar. EUR","% bekende kosten"]
    hc = list(range(2,9))
    ws.row_dimensions[10].height = 22
    for ci, h in zip(hc, headers):
        c = ws.cell(row=10, column=ci, value=h)
        c.fill = f(C["db"])
        c.font = ft(True,10,C["wh"])
        c.alignment = ac(True)
        c.border = bd()

    # Sorteer bekende kosten eerst, daarna onbekenden op volume
    def sort_key(item):
        sk, cnt = item
        pr = price_for_sku(sk)
        if pr is None:
            return (1, -cnt, sk)
        return (0, -(pr * cnt), sk)

    rows = sorted(lc.items(), key=sort_key)[:15]

    for ri, (sk, cnt) in enumerate(rows, 1):
        row = 10 + ri
        ws.row_dimensions[row].height = 18
        bg = C["ar"] if ri % 2 == 0 else C["wh"]

        pr = price_for_sku(sk)
        if pr is None:
            values = [name_for_sku(sk), sk, cnt, "Onbekend", "Onbekend", "Onbekend", ""]
            formats = [None, None, "#,##0", None, None, None, None]
        else:
            mo = pr * cnt
            yr = mo * 12
            pct = (mo / known_month_total) if known_month_total else 0
            values = [name_for_sku(sk), sk, cnt, pr, mo, yr, pct]
            formats = [None, None, "#,##0", "#,##0.00", "#,##0.00", "#,##0.00", "0.0%"]

        for ci, v, fmt in zip(hc, values, formats):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = f(bg if v != "Onbekend" else C["yl"])
            c.font = ft(s=10)
            c.alignment = al(True) if ci <= 3 else ac()
            c.border = bd()
            if fmt and isinstance(v, (int, float)):
                c.number_format = fmt

    tr = 10 + len(rows) + 1
    ws.row_dimensions[tr].height = 22
    total_values = ["TOTAAL (bekende prijzen)", "", sum(lc.values()), "", known_month_total, known_year_total, 1.0 if known_month_total else 0]
    total_formats = [None, None, "#,##0", None, "#,##0.00", "#,##0.00", "0.0%"]
    for ci, v, fmt in zip(hc, total_values, total_formats):
        c = ws.cell(row=tr, column=ci, value=v)
        c.fill = f(C["db"])
        c.font = ft(True,10,C["wh"])
        c.alignment = al() if ci <= 3 else ac()
        c.border = bd()
        if fmt and isinstance(v, (int, float)):
            c.number_format = fmt

    nr = tr + 2
    footer = (
        f"Opmerking: {unknown_assignment_count} licentietoewijzing(en) hebben geen actuele publieke NL-prijsreferentie "
        f"en worden daarom als 'Onbekend' getoond. Deze zijn niet meegenomen in de kosten-totalen."
    )
    ms(ws, f"B{nr}:H{nr}", footer, C["lg"], ft(8, i=True, c=C["dm"]), al(True))
    ws.row_dimensions[nr].height = 28

def license_overview(ws, lc, known_monthly_by_sku, known_month_total, known_year_total):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([2,32,22,14,16,18,18,18,2], 1):
        sw(ws, i, w)

    ms(ws, "B2:H2", "Licentie Overzicht  |  Alle SKU's bij Uitgeschakelde Accounts",
       C["mb"], ft(True,17,C["wh"]), ac(), 46)
    ws.row_dimensions[3].height = 4
    ms(ws, "B3:H3", "", C["ac"], ft(), ac())

    headers = ["Licentienaam","SKU Code","# Disabled","EUR p.u./mnd","Maand. EUR","Jaar. EUR","Besparing/jr EUR"]
    hc = list(range(2,9))
    ws.row_dimensions[5].height = 22
    for ci, h in zip(hc, headers):
        c = ws.cell(row=5, column=ci, value=h)
        c.fill = f(C["db"])
        c.font = ft(True,10,C["wh"])
        c.alignment = ac(True)
        c.border = bd()

    def sort_key(item):
        sk, cnt = item
        pr = price_for_sku(sk)
        if pr is None:
            return (1, -cnt, sk)
        return (0, -(pr * cnt), sk)

    rows = sorted(lc.items(), key=sort_key)
    for ri, (sk, cnt) in enumerate(rows, 1):
        row = 5 + ri
        ws.row_dimensions[row].height = 18
        bg = C["ar"] if ri % 2 == 0 else C["wh"]

        pr = price_for_sku(sk)
        if pr is None:
            values = [name_for_sku(sk), sk, cnt, "Onbekend", "Onbekend", "Onbekend", "Onbekend"]
            formats = [None, None, "#,##0", None, None, None, None]
        else:
            mo = pr * cnt
            yr = mo * 12
            values = [name_for_sku(sk), sk, cnt, pr, mo, yr, yr]
            formats = [None, None, "#,##0", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00"]

        for ci, v, fmt in zip(hc, values, formats):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = f(bg if v != "Onbekend" else C["yl"])
            c.font = ft(s=10)
            c.alignment = al(True) if ci <= 3 else ac()
            c.border = bd()
            if fmt and isinstance(v, (int, float)):
                c.number_format = fmt

    tr = 5 + len(rows) + 1
    ws.row_dimensions[tr].height = 24
    total_values = ["TOTAAL (bekende prijzen)", "", sum(lc.values()), "", known_month_total, known_year_total, known_year_total]
    total_formats = [None, None, "#,##0", None, "#,##0.00", "#,##0.00", "#,##0.00"]
    for ci, v, fmt in zip(hc, total_values, total_formats):
        c = ws.cell(row=tr, column=ci, value=v)
        c.fill = f(C["ac"])
        c.font = ft(True,11,C["wh"])
        c.alignment = al() if ci <= 3 else ac()
        c.border = bd()
        if fmt and isinstance(v, (int, float)):
            c.number_format = fmt

def user_details(ws, users):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([2,28,28,18,34,18,18,2], 1):
        sw(ws, i, w)

    ms(ws, "B2:G2", "Gebruikersdetail  |  Uitgeschakelde Accounts met Licenties",
       C["mb"], ft(True,16,C["wh"]), ac(), 44)
    ws.row_dimensions[3].height = 4
    ms(ws, "B3:G3", "", C["ac"], ft(), ac())

    headers = ["Weergavenaam","UserPrincipalName","Afd. / Functie","Licentie(s)","Maand. EUR","Jaar. EUR"]
    hc = list(range(2,8))
    ws.row_dimensions[5].height = 22
    for ci, h in zip(hc, headers):
        c = ws.cell(row=5, column=ci, value=h)
        c.fill = f(C["db"])
        c.font = ft(True,10,C["wh"])
        c.alignment = ac(True)
        c.border = bd()

    known_total = 0.0
    for ri, u in enumerate(users, 1):
        row = 5 + ri
        ws.row_dimensions[row].height = 18
        bg = C["ar"] if ri % 2 == 0 else C["wh"]

        lics = u.get("licenses") or []
        price_values = []
        lic_labels = []
        has_unknown = False

        for sk in lics:
            pr = price_for_sku(sk)
            nm = name_for_sku(sk)
            if pr is None:
                has_unknown = True
                lic_labels.append(f"{nm} [prijs onbekend]")
            else:
                price_values.append(pr)
                lic_labels.append(nm)

        month_known = sum(price_values)
        year_known = month_known * 12
        known_total += month_known

        if has_unknown:
            month_val = f"Min. EUR {month_known:,.2f} + onbekend"
            year_val = f"Min. EUR {year_known:,.2f} + onbekend"
        else:
            month_val = month_known
            year_val = year_known

        dept = (u.get("department") or "")
        jobt = (u.get("jobTitle") or "")
        dt = dept + (" / " + jobt if jobt else "")
        dn = u.get("displayName") or ""
        up = u.get("upn") or ""
        lns = "; ".join(lic_labels)

        values = [dn, up, dt, lns, month_val, year_val]
        formats = [None, None, None, None, "#,##0.00", "#,##0.00"]

        for ci, v, fmt in zip(hc, values, formats):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = f(bg if v != "Onbekend" else C["yl"])
            c.font = ft(s=9)
            c.alignment = al(True) if ci <= 5 else ac()
            c.border = bd()
            if fmt and isinstance(v, (int, float)):
                c.number_format = fmt

    tr = 5 + len(users) + 1
    ws.row_dimensions[tr].height = 22
    total_values = [f"TOTAAL bekende prijzen ({len(users)} gebruikers)", "", "", "", known_total, known_total * 12]
    total_formats = [None, None, None, None, "#,##0.00", "#,##0.00"]
    for ci, v, fmt in zip(hc, total_values, total_formats):
        c = ws.cell(row=tr, column=ci, value=v)
        c.fill = f(C["ac"])
        c.font = ft(True,10,C["wh"])
        c.alignment = al() if ci <= 5 else ac()
        c.border = bd()
        if fmt and isinstance(v, (int, float)):
            c.number_format = fmt

    ws.auto_filter.ref = "B5:G5"

def savings(ws, lc, known_month_total, known_year_total):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([2,30,14,18,18,18,22,2], 1):
        sw(ws, i, w)

    ms(ws, "B2:G2", "Besparingspotentieel  |  Wat betaalt u onnodig?",
       C["dg"], ft(True,18,C["wh"]), ac(), 50)
    ws.row_dimensions[3].height = 4
    ms(ws, "B3:G3", "", C["or"], ft(), ac())

    ws.row_dimensions[5].height = 50
    ws.row_dimensions[6].height = 20

    cards = [
        ("B5:C5", f"EUR {known_month_total:,.2f}", C["dg"], "B6:C6", "Bekend per maand"),
        ("D5:E5", f"EUR {known_year_total:,.2f}", C["or"], "D6:E6", "Bekend per jaar"),
        ("F5:G5", f"EUR {known_year_total:,.2f}", C["gr"], "F6:G6", "Max. bekende besparing/jr"),
    ]
    for top_rng, val, bg, bot_rng, lbl in cards:
        ws.merge_cells(top_rng)
        c = ws[top_rng.split(":")[0]]
        c.value = val
        c.fill = f(bg)
        c.font = ft(True,24,C["wh"])
        c.alignment = ac()

        ws.merge_cells(bot_rng)
        c2 = ws[bot_rng.split(":")[0]]
        c2.value = lbl
        c2.fill = f(bg)
        c2.font = ft(s=9, c=C["wh"])
        c2.alignment = ac()

    headers = ["Licentienaam","# Disabled","EUR p.u./mnd","Maand. EUR","Jaar. EUR","Besparing bij verwijdering EUR/jr"]
    hc = list(range(2,8))
    ws.row_dimensions[9].height = 22
    for ci, h in zip(hc, headers):
        c = ws.cell(row=9, column=ci, value=h)
        c.fill = f(C["db"])
        c.font = ft(True,10,C["wh"])
        c.alignment = ac(True)
        c.border = bd()

    def sort_key(item):
        sk, cnt = item
        pr = price_for_sku(sk)
        if pr is None:
            return (1, -cnt, sk)
        return (0, -(pr * cnt), sk)

    rows = sorted(lc.items(), key=sort_key)
    for ri, (sk, cnt) in enumerate(rows, 1):
        row = 9 + ri
        ws.row_dimensions[row].height = 18

        pr = price_for_sku(sk)
        if pr is None:
            values = [name_for_sku(sk), cnt, "Onbekend", "Onbekend", "Onbekend", "Onbekend"]
            formats = [None, "#,##0", None, None, None, None]
            bg = C["yl"]
        else:
            mo = pr * cnt
            yr = mo * 12
            sv = yr
            bg = C["yl"] if sv > 1000 else (C["ar"] if ri % 2 == 0 else C["wh"])
            values = [name_for_sku(sk), cnt, pr, mo, yr, sv]
            formats = [None, "#,##0", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00"]

        for ci, v, fmt in zip(hc, values, formats):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = f(bg)
            c.font = ft(s=10, b=(isinstance(v, (int, float)) and ci == 7 and v > 1000))
            c.alignment = al(True) if ci <= 2 else ac()
            c.border = bd()
            if fmt and isinstance(v, (int, float)):
                c.number_format = fmt

    tr = 9 + len(rows) + 1
    ws.row_dimensions[tr].height = 26
    total_values = ["TOTAAL MAX. BEKENDE BESPARING", sum(lc.values()), "", known_month_total, known_year_total, known_year_total]
    total_formats = [None, "#,##0", None, "#,##0.00", "#,##0.00", "#,##0.00"]
    for ci, v, fmt in zip(hc, total_values, total_formats):
        c = ws.cell(row=tr, column=ci, value=v)
        c.fill = f(C["gr"])
        c.font = ft(True,12,C["wh"])
        c.alignment = al() if ci <= 2 else ac()
        c.border = bd()
        if fmt and isinstance(v, (int, float)):
            c.number_format = fmt

def unknown_sheet(ws, lc):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([2,30,24,16,48,2], 1):
        sw(ws, i, w)

    ms(ws, "B2:E2", "Onbekende SKU's  |  Niet als gratis behandelen",
       C["or"], ft(True,16,C["wh"]), ac(), 44)
    ws.row_dimensions[3].height = 4
    ms(ws, "B3:E3", "", C["ac"], ft(), ac())

    headers = ["Licentienaam", "SKU Code", "# Toewijzingen", "Opmerking"]
    cols = [2,3,4,5]
    ws.row_dimensions[5].height = 22
    for ci, h in zip(cols, headers):
        c = ws.cell(row=5, column=ci, value=h)
        c.fill = f(C["db"])
        c.font = ft(True,10,C["wh"])
        c.alignment = ac(True)
        c.border = bd()

    unknown_rows = [(sk, cnt) for sk, cnt in lc.items() if price_for_sku(sk) is None]
    unknown_rows = sorted(unknown_rows, key=lambda x: (-x[1], x[0]))

    if not unknown_rows:
        c = ws.cell(row=6, column=2, value="Geen onbekende SKU's gevonden.")
        c.font = ft(True, 11)
        return

    for ri, (sk, cnt) in enumerate(unknown_rows, 1):
        row = 5 + ri
        bg = C["ar"] if ri % 2 == 0 else C["wh"]
        values = [
            name_for_sku(sk),
            sk,
            cnt,
            "Geen actuele publieke NL-prijsreferentie gevonden of legacy / tenant-specifieke SKU."
        ]
        for ci, v in zip(cols, values):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = f(bg)
            c.font = ft(s=10)
            c.alignment = al(True) if ci != 4 else ac()
            c.border = bd()
            if ci == 4 and isinstance(v, (int, float)):
                c.number_format = "#,##0"

def price_reference(ws):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([2,36,26,16,18,42,2], 1):
        sw(ws, i, w)

    ms(ws, "B2:F2", "Prijsreferentie  |  Publieke Microsoft NL referentieprijzen",
       C["mb"], ft(True,15,C["wh"]), ac(), 44)
    ws.row_dimensions[3].height = 4
    ms(ws, "B3:F3", "", C["ac"], ft(), ac())

    headers = ["Licentienaam", "SKU Code", "EUR p.u./mnd", "EUR p.u./jr", "Notitie"]
    cols = [2,3,4,5,6]
    ws.row_dimensions[5].height = 22
    for ci, h in zip(cols, headers):
        c = ws.cell(row=5, column=ci, value=h)
        c.fill = f(C["db"])
        c.font = ft(True,10,C["wh"])
        c.alignment = ac(True)
        c.border = bd()

    priced = [(sk, info) for sk, info in PRICE_TABLE.items() if info.get("price") is not None]
    priced = sorted(priced, key=lambda x: (-x[1]["price"], x[1]["name"]))

    for ri, (sk, info) in enumerate(priced, 1):
        row = 5 + ri
        bg = C["ar"] if ri % 2 == 0 else C["wh"]
        values = [
            info["name"],
            sk,
            info["price"],
            info["price"] * 12,
            "Publieke Microsoft NL referentieprijs; excl. btw; partnerafspraken kunnen afwijken."
        ]
        formats = [None, None, "#,##0.00", "#,##0.00", None]
        for ci, v, fmt in zip(cols, values, formats):
            c = ws.cell(row=row, column=ci, value=v)
            c.fill = f(bg)
            c.font = ft(s=9)
            c.alignment = al(True) if ci in (2,3,6) else ac()
            c.border = bd()
            if fmt and isinstance(v, (int, float)):
                c.number_format = fmt

    nr = 5 + len(priced) + 2
    footer = (
        "Belangrijk: onbekende, legacy, end-of-sale of tenant-specifieke SKU's worden niet als gratis behandeld. "
        "Die verschijnen elders in het rapport als 'Onbekend'."
    )
    ms(ws, f"B{nr}:F{nr}", footer, C["lg"], ft(s=8, i=True, c=C["dm"]), al(True))
    ws.row_dimensions[nr].height = 30

def run(data, out):
    wb = Workbook()
    wb.remove(wb.active)

    users = data.get("users", [])
    tenant = data.get("tenant", "Tenant")
    run_date = data.get("run_date", "")

    lc, unknown_skus, known_monthly_by_sku, known_month_total, known_year_total, unknown_assignment_count = build_summary(users)

    ws1 = wb.create_sheet("Dashboard")
    dashboard(ws1, tenant, run_date, users, lc, known_monthly_by_sku, known_month_total, known_year_total, unknown_assignment_count)

    ws2 = wb.create_sheet("Licentie Overzicht")
    license_overview(ws2, lc, known_monthly_by_sku, known_month_total, known_year_total)

    ws3 = wb.create_sheet("Gebruikers Detail")
    user_details(ws3, users)

    ws4 = wb.create_sheet("Besparingspotentieel")
    savings(ws4, lc, known_month_total, known_year_total)

    ws5 = wb.create_sheet("Onbekende SKU's")
    unknown_sheet(ws5, lc)

    ws6 = wb.create_sheet("Prijsreferentie")
    price_reference(ws6)

    wb.save(out)
    print("[OK] Rapport opgeslagen: " + out)

if __name__ == "__main__":
    with open(sys.argv[1], "r", encoding="utf-8") as fh:
        d = json.load(fh)
    run(d, sys.argv[2])
'@

$pyCode | Set-Content -Path $tempPy -Encoding UTF8

Write-Host "     Excel-opmaak en prijsreferenties verwerken..." -ForegroundColor Gray

try {
    $result = & $pythonCmd $tempPy $tempJson $outputFull 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Python exitcode $LASTEXITCODE : $result"
    }
    Write-Host "     $result" -ForegroundColor Green
} catch {
    Write-Host "  FOUT bij Excel-generatie: $_" -ForegroundColor Red
    exit 1
} finally {
    Remove-Item $tempJson -Force -ErrorAction SilentlyContinue
    Remove-Item $tempPy   -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "  =================================================================" -ForegroundColor Green
Write-Host "    RAPPORT SUCCESVOL GEGENEREERD" -ForegroundColor Green
Write-Host "  =================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Bestand    : $outputFull" -ForegroundColor White
Write-Host "  Tenant     : $tenantName" -ForegroundColor White
Write-Host "  Gebruikers : $($allUsers.Count) uitgeschakelde accounts met licentie(s)" -ForegroundColor Yellow

$totLic = ($allUsers | ForEach-Object { @($_.licenses).Count } | Measure-Object -Sum).Sum
Write-Host "  Licenties  : $totLic gekoppelde licenties (totaal)" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Tabbladen in het rapport:" -ForegroundColor Cyan
Write-Host "     Dashboard             - KPI overzicht en toplicenties" -ForegroundColor Gray
Write-Host "     Licentie Overzicht    - Alle SKU's met kostenstatus" -ForegroundColor Gray
Write-Host "     Gebruikers Detail     - Volledig gebruikersoverzicht" -ForegroundColor Gray
Write-Host "     Besparingspotentieel  - Kostenanalyse" -ForegroundColor Gray
Write-Host "     Onbekende SKU's       - Niet-publiek geprijsde / legacy SKU's" -ForegroundColor Gray
Write-Host "     Prijsreferentie       - Publieke Microsoft NL referentieprijzen" -ForegroundColor Gray
Write-Host ""

$openFile = Read-Host "  Wilt u het rapport nu openen? (J/N)"
if ($openFile -match "^[Jj]") {
    try {
        Start-Process $outputFull
        Write-Host "  Rapport geopend." -ForegroundColor Green
    } catch {
        Write-Host "  Open handmatig: $outputFull" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "  Gereed." -ForegroundColor DarkGray
Write-Host ""