<#
.SYNOPSIS
    Exporteert alle Conditional Access policies naar een Excel rapport via Microsoft Graph API.

.DESCRIPTION
    Dit script maakt verbinding met Microsoft Graph API en exporteert alle Conditional Access policies
    naar een geformatteerd Excel rapport (.xlsx), geschikt voor maandelijkse service level rapportages.
    Het rapport bevat: policy naam, status, conditie-details (gebruikers, apps, locaties),
    grant/session controls en aanmaak- en wijzigingsdatum.
    
    Het script ondersteunt zowel delegated (interactief) als application (service principal) authenticatie.

.PARAMETER OutputPath
    Het pad waar het Excel bestand moet worden opgeslagen.
    Standaard: .\ConditionalAccessPolicies_Rapport_<datum>.xlsx

.PARAMETER TenantId
    De Tenant ID van je Azure AD tenant (optioneel bij delegated auth)

.PARAMETER ClientId
    De Client ID van je Azure AD App Registration (optioneel)

.PARAMETER UseServicePrincipal
    Switch om authenticatie via service principal te gebruiken (vereist ClientId, TenantId en ClientSecret)

.PARAMETER ClientSecret
    Het client secret voor service principal authenticatie

.EXAMPLE
    .\Export-ConditionalAccessPolicies.ps1
    
.EXAMPLE
    .\Export-ConditionalAccessPolicies.ps1 -OutputPath "C:\Rapporten\CA-Policies.xlsx"

.EXAMPLE
    .\Export-ConditionalAccessPolicies.ps1 -TenantId "your-tenant-id" -ClientId "your-client-id" -UseServicePrincipal -ClientSecret "your-secret"

.NOTES
    Vereiste permissies:
    - Policy.Read.All        (Conditional Access policies)
    - User.Read.All          (gebruikersnamen opzoeken)
    - Group.Read.All         (groepsnamen opzoeken)
    - Application.Read.All   (applicatienamen opzoeken)
    - Directory.Read.All     (directory rollen opzoeken)
    
    Voor delegated (interactief): User moet Global Reader of Conditional Access Administrator zijn
    Voor application: App Registration moet bovenstaande API permissions hebben

    Vereiste software:
    - PowerShell 5.1 of hoger
    - Microsoft.Graph modules (worden automatisch geïnstalleerd indien nodig)
    - ImportExcel module (wordt automatisch geïnstalleerd indien nodig)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\ConditionalAccessPolicies_Rapport_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [switch]$UseServicePrincipal,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret
)

#region ── Module beheer ────────────────────────────────────────────────────────

function Install-RequiredModules {
    Write-Host "Controleren van vereiste modules..." -ForegroundColor Cyan

    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Applications',
        'ImportExcel'
    )

    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Host "  Module '$module' niet gevonden. Installeren..." -ForegroundColor Yellow
            try {
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Host "  Module '$module' succesvol geïnstalleerd." -ForegroundColor Green
            }
            catch {
                Write-Error "Fout bij installeren van module '$module': $_"
                return $false
            }
        }
        else {
            Write-Host "  Module '$module' is beschikbaar." -ForegroundColor Green
        }
    }
    return $true
}

#endregion

#region ── Authenticatie ────────────────────────────────────────────────────────

function Connect-ToGraph {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [switch]$UseServicePrincipal,
        [string]$ClientSecret
    )

    Write-Host "`nVerbinden met Microsoft Graph..." -ForegroundColor Cyan

    try {
        if ($UseServicePrincipal) {
            if ([string]::IsNullOrEmpty($ClientSecret)) {
                Write-Error "ClientSecret is vereist voor service principal authenticatie."
                return $false
            }
            $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
            $credential   = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
        }
        else {
            $connectParams = @{
                Scopes    = @(
                    'Policy.Read.All',
                    'User.Read.All',
                    'Group.Read.All',
                    'Application.Read.All',
                    'Directory.Read.All'
                )
                NoWelcome = $true
            }
            if (![string]::IsNullOrEmpty($TenantId)) { $connectParams['TenantId']  = $TenantId }
            if (![string]::IsNullOrEmpty($ClientId))  { $connectParams['ClientId']  = $ClientId  }
            Connect-MgGraph @connectParams
        }

        $ctx = Get-MgContext
        Write-Host "  Verbonden met tenant : $($ctx.TenantId)" -ForegroundColor Green
        Write-Host "  Account              : $($ctx.Account)"  -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Fout bij verbinden met Microsoft Graph: $_"
        return $false
    }
}

#endregion

#region ── Data ophalen ─────────────────────────────────────────────────────────

function Get-ConditionalAccessPolicies {
    Write-Host "`nOphalen van Conditional Access policies..." -ForegroundColor Cyan

    try {
        $policies = Get-MgIdentityConditionalAccessPolicy -All

        if ($null -eq $policies -or $policies.Count -eq 0) {
            Write-Warning "Geen Conditional Access policies gevonden."
            return @()
        }

        $sorted = $policies | Sort-Object DisplayName

        # ── Console-overzicht met kleurcodering ──────────────────────────────
        Write-Host "`n  Gevonden policies ($($sorted.Count)):" -ForegroundColor Cyan
        Write-Host "  $('─' * 70)" -ForegroundColor DarkGray

        foreach ($p in $sorted) {
            $kleur = switch ($p.State) {
                "enabled"                          { "Green"      }
                "disabled"                         { "Yellow"     }
                "enabledForReportingButNotEnforced" { "DarkYellow" }
                default                            { "Gray"       }
            }
            $label = switch ($p.State) {
                "enabled"                          { "ACTIEF"      }
                "disabled"                         { "UITGESCHAKELD" }
                "enabledForReportingButNotEnforced" { "ALLEEN RAPPORT" }
                default                            { $p.State      }
            }

            Write-Host "  ● " -NoNewline -ForegroundColor $kleur
            Write-Host "$($p.DisplayName.PadRight(55))" -NoNewline
            Write-Host "[$label]" -ForegroundColor $kleur
        }
        Write-Host "  $('─' * 70)" -ForegroundColor DarkGray

        return $sorted
    }
    catch {
        Write-Error "Fout bij ophalen van policies: $_"
        return @()
    }
}

#endregion

#region ── Naam-lookup cache & resolvers ────────────────────────────────────────

# Globale caches — worden eenmalig gevuld, daarna hergebruikt
$script:CacheGebruikers  = @{}
$script:CacheGroepen     = @{}
$script:CacheApps        = @{}
$script:CacheLocaties    = @{}
$script:CacheRollen      = @{}

# Bekende GUIDs die nooit via de API worden opgezocht
$script:VasteWaarden = @{
    # Gebruikers / groepen
    'All'                                   = 'Alle gebruikers'
    'GuestsOrExternalUsers'                 = 'Gasten en externe gebruikers'
    'None'                                  = 'Geen'
    # Toepassingen
    'AllApplications'                       = 'Alle toepassingen'
    'Office365'                             = 'Office 365'
    # Locaties
    'AllTrusted'                            = 'Alle vertrouwde locaties'
    'AllCompliantNetworkLocations'          = 'Alle compatibele netwerklocaties'
}

# Bekende Microsoft service-app GUIDs → leesbare naam
$script:BekendeApps = @{
    '00000002-0000-0ff1-ce00-000000000000' = 'Office 365 Exchange Online'
    '00000003-0000-0ff1-ce00-000000000000' = 'Office 365 SharePoint Online'
    '00000003-0000-0000-c000-000000000000' = 'Microsoft Graph'
    '00000004-0000-0ff1-ce00-000000000000' = 'Skype for Business Online'
    '00000005-0000-0ff1-ce00-000000000000' = 'Microsoft Flow'
    '00000006-0000-0ff1-ce00-000000000000' = 'Microsoft Office'
    '00000007-0000-0ff1-ce00-000000000000' = 'Microsoft Azure Active Directory'
    '00000009-0000-0000-c000-000000000000' = 'Microsoft Power BI'
    '0000000a-0000-0000-c000-000000000000' = 'Microsoft App Access Panel'
    '0000000c-0000-0000-c000-000000000000' = 'Microsoft App (Azure Portal)'
    '00000012-0000-0000-c000-000000000000' = 'Microsoft Rights Management'
    '00000014-0000-0000-c000-000000000000' = 'Microsoft Intune'
    '1fec8e78-bce4-4aaf-ab1b-5451cc387264' = 'Microsoft Teams'
    '4765445b-32c6-49b0-83e6-1d93765276ca' = 'Office Online Core SSO'
    '57fb890c-0dab-4253-a5e0-7188c88b2bb4' = 'SharePoint Online Web Client Extensibility'
    '89bee1f7-5e6e-4d8a-9f3d-ecd601259da7' = 'Office Online Search'
    'd3590ed6-52b3-4102-aeff-aad2292ab01c' = 'Microsoft Office (desktop)'
    'Office365'                             = 'Office 365 (alle apps)'
}

function Initialize-LookupCaches {
    Write-Host "`nOphalen van lookup-gegevens (gebruikers, groepen, apps, locaties)..." -ForegroundColor Cyan

    # ── Alle named locations ophalen ─────────────────────────────────────────
    try {
        $locations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
        foreach ($loc in $locations) {
            $script:CacheLocaties[$loc.Id] = $loc.DisplayName
        }
        Write-Host "  Locaties geladen     : $($script:CacheLocaties.Count)" -ForegroundColor Green
    }
    catch {
        Write-Warning "  Kon locaties niet ophalen: $_"
    }

    # ── Alle directory role templates ophalen (voor rol-GUIDs) ────────────────
    try {
        $roleTemplates = Get-MgDirectoryRoleTemplate -All -ErrorAction Stop
        foreach ($rt in $roleTemplates) {
            $script:CacheRollen[$rt.Id] = $rt.DisplayName
        }
        Write-Host "  Directoryrol templates: $($script:CacheRollen.Count)" -ForegroundColor Green
    }
    catch {
        Write-Warning "  Kon rol templates niet ophalen: $_"
    }
}

# Hulpfunctie: bepaal of een string een GUID is
function Test-IsGuid {
    param([string]$value)
    return ($value -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
}

# Resolve gebruiker GUID → displayName
function Resolve-GebruikerNaam {
    param([string]$id)
    if ([string]::IsNullOrEmpty($id)) { return $id }
    if ($script:VasteWaarden.ContainsKey($id)) { return $script:VasteWaarden[$id] }
    if (!(Test-IsGuid $id)) { return $id }
    if ($script:CacheGebruikers.ContainsKey($id)) { return $script:CacheGebruikers[$id] }

    try {
        $user = Get-MgUser -UserId $id -Property DisplayName -ErrorAction Stop
        $naam = $user.DisplayName
    }
    catch {
        # Mogelijk een groep of service principal — probeer groep
        $naam = Resolve-GroepNaam $id
    }
    $script:CacheGebruikers[$id] = $naam
    return $naam
}

# Resolve groep GUID → displayName
function Resolve-GroepNaam {
    param([string]$id)
    if ([string]::IsNullOrEmpty($id)) { return $id }
    if ($script:VasteWaarden.ContainsKey($id)) { return $script:VasteWaarden[$id] }
    if (!(Test-IsGuid $id)) { return $id }
    if ($script:CacheGroepen.ContainsKey($id)) { return $script:CacheGroepen[$id] }

    try {
        $group = Get-MgGroup -GroupId $id -Property DisplayName -ErrorAction Stop
        $naam  = $group.DisplayName
    }
    catch {
        $naam = $id   # Fallback: toon ID als naam niet gevonden
    }
    $script:CacheGroepen[$id] = $naam
    return $naam
}

# Resolve applicatie GUID/keyword → displayName
function Resolve-AppNaam {
    param([string]$id)
    if ([string]::IsNullOrEmpty($id)) { return $id }
    if ($script:BekendeApps.ContainsKey($id))  { return $script:BekendeApps[$id]  }
    if ($script:VasteWaarden.ContainsKey($id)) { return $script:VasteWaarden[$id] }
    if (!(Test-IsGuid $id)) { return $id }
    if ($script:CacheApps.ContainsKey($id)) { return $script:CacheApps[$id] }

    try {
        # Probeer eerst als Service Principal (AppId)
        $sp = Get-MgServicePrincipal -Filter "appId eq '$id'" -Property DisplayName -ErrorAction Stop
        if ($sp) {
            $naam = $sp.DisplayName
        }
        else {
            # Probeer als Object ID
            $sp2  = Get-MgServicePrincipal -ServicePrincipalId $id -Property DisplayName -ErrorAction Stop
            $naam = $sp2.DisplayName
        }
    }
    catch {
        $naam = $id   # Fallback
    }
    $script:CacheApps[$id] = $naam
    return $naam
}

# Resolve locatie GUID/keyword → displayName
function Resolve-LocatieNaam {
    param([string]$id)
    if ([string]::IsNullOrEmpty($id)) { return $id }
    if ($script:VasteWaarden.ContainsKey($id)) { return $script:VasteWaarden[$id] }
    if ($script:CacheLocaties.ContainsKey($id)) { return $script:CacheLocaties[$id] }
    if (!(Test-IsGuid $id)) { return $id }
    return $id   # Onbekende locatie-GUID: toon als-is
}

# Resolve rol GUID → displayName
function Resolve-RolNaam {
    param([string]$id)
    if ([string]::IsNullOrEmpty($id)) { return $id }
    if ($script:CacheRollen.ContainsKey($id)) { return $script:CacheRollen[$id] }
    if (!(Test-IsGuid $id)) { return $id }
    return $id   # Onbekende rol-GUID
}

# Hulpfunctie: lijst van IDs resolven naar namen, kommagescheiden
function Resolve-Lijst {
    param(
        [string[]]$ids,
        [scriptblock]$resolver
    )
    if ($null -eq $ids -or $ids.Count -eq 0) { return $null }
    $namen = $ids | ForEach-Object { & $resolver $_ }
    return ($namen | Where-Object { $_ }) -join ", "
}

#endregion

#region ── Helperfuncties voor leesbaarheid ─────────────────────────────────────

function Format-List {
    param([object]$value)
    if ($null -eq $value) { return "—" }
    $list = @($value)
    if ($list.Count -eq 0) { return "—" }
    return ($list | Where-Object { $_ }) -join ", "
}

function Format-Users {
    param([object]$conditions)
    $parts = @()

    $includeUsers  = @($conditions.Users.IncludeUsers  | Where-Object { $_ })
    $excludeUsers  = @($conditions.Users.ExcludeUsers  | Where-Object { $_ })
    $includeGroups = @($conditions.Users.IncludeGroups | Where-Object { $_ })
    $excludeGroups = @($conditions.Users.ExcludeGroups | Where-Object { $_ })
    $includeRoles  = @($conditions.Users.IncludeRoles  | Where-Object { $_ })
    $excludeRoles  = @($conditions.Users.ExcludeRoles  | Where-Object { $_ })

    if ($includeUsers.Count -gt 0) {
        $namen = Resolve-Lijst -ids $includeUsers -resolver { param($id) Resolve-GebruikerNaam $id }
        $parts += "Incl. gebruikers: $namen"
    }
    if ($excludeUsers.Count -gt 0) {
        $namen = Resolve-Lijst -ids $excludeUsers -resolver { param($id) Resolve-GebruikerNaam $id }
        $parts += "Excl. gebruikers: $namen"
    }
    if ($includeGroups.Count -gt 0) {
        $namen = Resolve-Lijst -ids $includeGroups -resolver { param($id) Resolve-GroepNaam $id }
        $parts += "Incl. groepen: $namen"
    }
    if ($excludeGroups.Count -gt 0) {
        $namen = Resolve-Lijst -ids $excludeGroups -resolver { param($id) Resolve-GroepNaam $id }
        $parts += "Excl. groepen: $namen"
    }
    if ($includeRoles.Count -gt 0) {
        $namen = Resolve-Lijst -ids $includeRoles -resolver { param($id) Resolve-RolNaam $id }
        $parts += "Incl. rollen: $namen"
    }
    if ($excludeRoles.Count -gt 0) {
        $namen = Resolve-Lijst -ids $excludeRoles -resolver { param($id) Resolve-RolNaam $id }
        $parts += "Excl. rollen: $namen"
    }

    if ($parts.Count -eq 0) { return "—" }
    return $parts -join " | "
}

function Format-Apps {
    param([object]$conditions)
    $parts = @()

    $includeApps = @($conditions.Applications.IncludeApplications | Where-Object { $_ })
    $excludeApps = @($conditions.Applications.ExcludeApplications | Where-Object { $_ })
    $userActions = @($conditions.Applications.IncludeUserActions   | Where-Object { $_ })

    if ($includeApps.Count -gt 0) {
        $namen = Resolve-Lijst -ids $includeApps -resolver { param($id) Resolve-AppNaam $id }
        $parts += "Incl.: $namen"
    }
    if ($excludeApps.Count -gt 0) {
        $namen = Resolve-Lijst -ids $excludeApps -resolver { param($id) Resolve-AppNaam $id }
        $parts += "Excl.: $namen"
    }
    if ($userActions.Count -gt 0) {
        $parts += "Gebruikersacties: $(Format-List $userActions)"
    }

    if ($parts.Count -eq 0) { return "—" }
    return $parts -join " | "
}

function Format-Locations {
    param([object]$conditions)
    $parts = @()

    $includeLocations = @($conditions.Locations.IncludeLocations | Where-Object { $_ })
    $excludeLocations = @($conditions.Locations.ExcludeLocations | Where-Object { $_ })

    if ($includeLocations.Count -gt 0) {
        $namen = Resolve-Lijst -ids $includeLocations -resolver { param($id) Resolve-LocatieNaam $id }
        $parts += "Incl.: $namen"
    }
    if ($excludeLocations.Count -gt 0) {
        $namen = Resolve-Lijst -ids $excludeLocations -resolver { param($id) Resolve-LocatieNaam $id }
        $parts += "Excl.: $namen"
    }

    if ($parts.Count -eq 0) { return "—" }
    return $parts -join " | "
}

function Format-GrantControls {
    param([object]$grant)
    if ($null -eq $grant) { return "—" }

    $controls   = Format-List $grant.BuiltInControls
    $operator   = $grant.Operator
    $customAuth = Format-List $grant.CustomAuthenticationFactors
    $termsOfUse = Format-List $grant.TermsOfUse

    $parts = @()
    if ($controls   -ne "—") { $parts += "Controls ($operator): $controls" }
    if ($customAuth -ne "—") { $parts += "Custom auth: $customAuth"         }
    if ($termsOfUse -ne "—") { $parts += "Gebruiksvoorwaarden: $termsOfUse" }

    if ($parts.Count -eq 0) { return "—" }
    return $parts -join " | "
}

function Format-SessionControls {
    param([object]$session)
    if ($null -eq $session) { return "—" }

    $parts = @()

    if ($session.ApplicationEnforcedRestrictions.IsEnabled) {
        $parts += "App-beperkingen"
    }
    if ($null -ne $session.CloudAppSecurity -and $session.CloudAppSecurity.IsEnabled) {
        $parts += "Cloud App Security: $($session.CloudAppSecurity.CloudAppSecurityType)"
    }
    if ($null -ne $session.SignInFrequency -and $session.SignInFrequency.IsEnabled) {
        $parts += "Aanmeldfrequentie: $($session.SignInFrequency.Value) $($session.SignInFrequency.Type)"
    }
    if ($null -ne $session.PersistentBrowser -and $session.PersistentBrowser.IsEnabled) {
        $parts += "Persistente browser: $($session.PersistentBrowser.Mode)"
    }
    if ($null -ne $session.DisableResilienceDefaults) {
        $parts += "Resilience uitgeschakeld: $($session.DisableResilienceDefaults)"
    }

    if ($parts.Count -eq 0) { return "—" }
    return $parts -join " | "
}

function Get-StatusLabel {
    param([string]$state)
    switch ($state) {
        "enabled"                          { return "Actief"           }
        "disabled"                         { return "Uitgeschakeld"    }
        "enabledForReportingButNotEnforced" { return "Alleen rapport"   }
        default                            { return $state             }
    }
}

#endregion

#region ── Excel export ─────────────────────────────────────────────────────────

function Export-NaarExcel {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Policies,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Host "`nExcel rapport aanmaken..." -ForegroundColor Cyan

    # ── Directory aanmaken indien nodig ──────────────────────────────────────
    $dir = Split-Path -Path $OutputPath -Parent
    if (![string]::IsNullOrEmpty($dir) -and !(Test-Path -Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $ctx        = Get-MgContext
    $exportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $maand      = Get-Date -Format "MMMM yyyy"

    # ── Rijdata opbouwen ─────────────────────────────────────────────────────
    $rows = foreach ($p in $Policies) {
        [PSCustomObject]@{
            'Naam'                  = $p.DisplayName
            'Status'                = Get-StatusLabel $p.State
            'Gebruikers / Groepen'  = Format-Users     $p.Conditions
            'Toepassingen'          = Format-Apps      $p.Conditions
            'Locaties'              = Format-Locations  $p.Conditions
            'Grant Controls'        = Format-GrantControls  $p.GrantControls
            'Session Controls'      = Format-SessionControls $p.SessionControls
            'Aangemaakt op'         = if ($p.CreatedDateTime)  { ([datetime]$p.CreatedDateTime).ToString("dd-MM-yyyy HH:mm")  } else { "—" }
            'Gewijzigd op'          = if ($p.ModifiedDateTime) { ([datetime]$p.ModifiedDateTime).ToString("dd-MM-yyyy HH:mm") } else { "—" }
        }
    }

    try {
        # ── Excel pakket opbouwen met ImportExcel ────────────────────────────
        $excel = $rows | Export-Excel -Path $OutputPath `
            -WorksheetName "CA Policies" `
            -AutoSize `
            -FreezeTopRow `
            -BoldTopRow `
            -PassThru

        $ws = $excel.Workbook.Worksheets["CA Policies"]

        # ── Stijlen definiëren ───────────────────────────────────────────────
        Add-Type -AssemblyName System.Drawing
        $kleurActief      = [System.Drawing.Color]::FromArgb(198, 239, 206)   # lichtgroen
        $kleurUitge       = [System.Drawing.Color]::FromArgb(255, 235, 156)   # lichtgeel
        $kleurRapport     = [System.Drawing.Color]::FromArgb(255, 199, 134)   # lichteoranje
        $kleurHeader      = [System.Drawing.Color]::FromArgb(31,  78,  121)   # donkerblauw
        $kleurHeaderTekst = [System.Drawing.Color]::White
        $kleurRijOneven   = [System.Drawing.Color]::FromArgb(242, 242, 242)   # lichtgrijs

        # Kolomindex van 'Status' (kolom B = 2)
        $statusKolom = 2

        # ── Headerrij opmaken ────────────────────────────────────────────────
        $aantalKolommen = $ws.Dimension.Columns
        for ($k = 1; $k -le $aantalKolommen; $k++) {
            $cel = $ws.Cells[1, $k]
            $cel.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cel.Style.Fill.BackgroundColor.SetColor($kleurHeader)
            $cel.Style.Font.Color.SetColor($kleurHeaderTekst)
            $cel.Style.Font.Bold = $true
            $cel.Style.Font.Name = "Arial"
            $cel.Style.Font.Size = 10
            $cel.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        }

        # ── Datarijen opmaken ────────────────────────────────────────────────
        $aantalRijen = $ws.Dimension.Rows

        for ($r = 2; $r -le $aantalRijen; $r++) {
            $statusWaarde = $ws.Cells[$r, $statusKolom].Value

            # Rijachtergrond bepalen op basis van status
            $rijKleur = switch ($statusWaarde) {
                "Actief"         { $kleurActief  }
                "Uitgeschakeld"  { $kleurUitge   }
                "Alleen rapport" { $kleurRapport }
                default          { if ($r % 2 -eq 0) { $kleurRijOneven } else { [System.Drawing.Color]::White } }
            }

            for ($k = 1; $k -le $aantalKolommen; $k++) {
                $cel = $ws.Cells[$r, $k]

                # Achtergrondkleur
                $cel.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $cel.Style.Fill.BackgroundColor.SetColor($rijKleur)

                # Lettertype
                $cel.Style.Font.Name = "Arial"
                $cel.Style.Font.Size = 9

                # Tekstafbreking en uitlijning
                $cel.Style.WrapText = $true
                $cel.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
            }

            # Status-cel: vetgedrukt
            $ws.Cells[$r, $statusKolom].Style.Font.Bold = $true
            $ws.Cells[$r, $statusKolom].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        }

        # ── Kolombreedtes instellen ──────────────────────────────────────────
        $ws.Column(1).Width = 40   # Naam
        $ws.Column(2).Width = 18   # Status
        $ws.Column(3).Width = 50   # Gebruikers / Groepen
        $ws.Column(4).Width = 45   # Toepassingen
        $ws.Column(5).Width = 35   # Locaties
        $ws.Column(6).Width = 45   # Grant Controls
        $ws.Column(7).Width = 45   # Session Controls
        $ws.Column(8).Width = 20   # Aangemaakt op
        $ws.Column(9).Width = 20   # Gewijzigd op

        # ── Randen toevoegen ─────────────────────────────────────────────────
        $bereik = $ws.Cells[1, 1, $aantalRijen, $aantalKolommen]
        $bereik.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $bereik.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $bereik.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $bereik.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

        # ── Legenda-tabblad toevoegen ────────────────────────────────────────
        $wsLeg = $excel.Workbook.Worksheets.Add("Legenda & Info")

        $legenda = @(
            @{ Rij = 1;  Tekst = "Conditional Access Policy Rapport"; Vet = $true;  Grootte = 14 }
            @{ Rij = 2;  Tekst = "Maandelijkse Service Level Rapportage — $maand";   Vet = $false; Grootte = 11 }
            @{ Rij = 3;  Tekst = "";                                                  Vet = $false; Grootte = 10 }
            @{ Rij = 4;  Tekst = "Exportdatum       : $exportDate";                  Vet = $false; Grootte = 10 }
            @{ Rij = 5;  Tekst = "Geëxporteerd door : $($env:USERNAME)";             Vet = $false; Grootte = 10 }
            @{ Rij = 6;  Tekst = "Tenant ID         : $($ctx.TenantId)";             Vet = $false; Grootte = 10 }
            @{ Rij = 7;  Tekst = "Aantal policies   : $($Policies.Count)";           Vet = $false; Grootte = 10 }
            @{ Rij = 8;  Tekst = "";                                                  Vet = $false; Grootte = 10 }
            @{ Rij = 9;  Tekst = "Opgeloste object-namen (via Graph API):";          Vet = $true;  Grootte = 10 }
            @{ Rij = 10; Tekst = "  Gebruikers  : $($script:CacheGebruikers.Count)"; Vet = $false; Grootte = 10 }
            @{ Rij = 11; Tekst = "  Groepen     : $($script:CacheGroepen.Count)";    Vet = $false; Grootte = 10 }
            @{ Rij = 12; Tekst = "  Toepassingen: $($script:CacheApps.Count)";       Vet = $false; Grootte = 10 }
            @{ Rij = 13; Tekst = "  Locaties    : $($script:CacheLocaties.Count)";   Vet = $false; Grootte = 10 }
            @{ Rij = 14; Tekst = "";                                                  Vet = $false; Grootte = 10 }
            @{ Rij = 15; Tekst = "Kleurlegenda";                                     Vet = $true;  Grootte = 11 }
        )

        foreach ($item in $legenda) {
            $c = $wsLeg.Cells[$item.Rij, 1]
            $c.Value          = $item.Tekst
            $c.Style.Font.Name = "Arial"
            $c.Style.Font.Size = $item.Grootte
            $c.Style.Font.Bold = $item.Vet
        }

        # Kleurlegenda blokjes
        $legItems = @(
            @{ Rij = 16; Label = "Actief";          Kleur = $kleurActief  }
            @{ Rij = 17; Label = "Uitgeschakeld";   Kleur = $kleurUitge   }
            @{ Rij = 18; Label = "Alleen rapport";  Kleur = $kleurRapport }
        )

        foreach ($li in $legItems) {
            $cKleur = $wsLeg.Cells[$li.Rij, 1]
            $cKleur.Value = "   "
            $cKleur.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cKleur.Style.Fill.BackgroundColor.SetColor($li.Kleur)
            $cKleur.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $cKleur.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $cKleur.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $cKleur.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

            $cLabel = $wsLeg.Cells[$li.Rij, 2]
            $cLabel.Value          = $li.Label
            $cLabel.Style.Font.Name = "Arial"
            $cLabel.Style.Font.Size = 10
        }

        $wsLeg.Column(1).Width = 6
        $wsLeg.Column(2).Width = 25

        # Zet CA Policies tabblad als actief
        $excel.Workbook.View.ActiveTab = 0

        # ── Opslaan ──────────────────────────────────────────────────────────
        $excel.Save()
        $excel.Dispose()

        $grootte = [math]::Round((Get-Item $OutputPath).Length / 1KB, 2)
        Write-Host "  Rapport opgeslagen : $OutputPath" -ForegroundColor Green
        Write-Host "  Bestandsgrootte    : $grootte KB"  -ForegroundColor Green

        return $true
    }
    catch {
        Write-Error "Fout bij aanmaken van Excel rapport: $_"
        return $false
    }
}

#endregion

#region ── Main ─────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   Conditional Access Policy Rapport Tool     ║" -ForegroundColor Cyan
Write-Host "║   Maandelijkse Service Level Rapportage       ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Stap 1: Modules
if (!(Install-RequiredModules)) {
    Write-Error "Kan niet doorgaan zonder vereiste modules."
    exit 1
}

# Stap 2: Verbinding
$connectParams = @{}
if (![string]::IsNullOrEmpty($TenantId))    { $connectParams['TenantId']           = $TenantId    }
if (![string]::IsNullOrEmpty($ClientId))    { $connectParams['ClientId']            = $ClientId    }
if ($UseServicePrincipal)                   { $connectParams['UseServicePrincipal'] = $true
                                              $connectParams['ClientSecret']        = $ClientSecret }

if (!(Connect-ToGraph @connectParams)) {
    Write-Error "Kan niet verbinden met Microsoft Graph."
    exit 1
}

# Stap 3: Policies ophalen
$policies = Get-ConditionalAccessPolicies

if ($policies.Count -eq 0) {
    Write-Warning "Geen policies gevonden om te exporteren."
    Disconnect-MgGraph | Out-Null
    exit 0
}

# Stap 4: Lookup-caches opbouwen (namen voor IDs)
Initialize-LookupCaches

# Stap 5: Excel rapport genereren
if (Export-NaarExcel -Policies $policies -OutputPath $OutputPath) {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   ✔  Rapport succesvol aangemaakt!           ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Samenvatting:" -ForegroundColor Cyan
    Write-Host "  ● Aantal policies : $($policies.Count)"   -ForegroundColor White
    Write-Host "  ● Opgeslagen in   : $OutputPath"           -ForegroundColor White
    Write-Host ""
}
else {
    Write-Error "Export mislukt."
    Disconnect-MgGraph | Out-Null
    exit 1
}

# Stap 6: Disconnect
Write-Host "Verbinding verbreken..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Klaar!" -ForegroundColor Green
Write-Host ""

#endregion
